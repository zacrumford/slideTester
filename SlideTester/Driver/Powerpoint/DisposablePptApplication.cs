using System;
using System.Diagnostics;
using System.Threading;
using System.Threading.Tasks;

using SlideTester.Common;
using SlideTester.Common.Log;
using PptApplication = Microsoft.Office.Interop.PowerPoint.Application;
using PptPresentation = Microsoft.Office.Interop.PowerPoint.Presentation;
using PptSlide = Microsoft.Office.Interop.PowerPoint.Slide;
using PanoptoSlide = SlideTester.Driver.Slide;

namespace SlideTester.Driver.Powerpoint
{
    /// <summary>
    /// Helper method to wrap the Microsoft.Office.PptApplication object as a disposable object.
    /// </summary>
    /// <remarks>
    /// We can certainly do more to make this better but we invested minimal work for now. If we decide to
    /// Kill the PowerpointSlideDriver then this class will die. If we decide to keep PowerpointSlideDriver then
    /// we may want to clean this up more 
    /// </remarks>
    internal class DisposablePptApplication : SafeDisposable
    {
        #region Members and properties

        /// <summary>
        /// Handler to ppt application which will need cleaning up upon dispose
        /// </summary>
        public PptApplication ApplicationHandle { get; private set; }
        
        #endregion

        /// <summary>
        /// Initial values ctor
        /// </summary>
        public DisposablePptApplication(
            PptApplication applicationHandle)
        {
            this.ApplicationHandle = applicationHandle;
        }

        /// <summary>
        /// Method to clean up managed resources.  This method will be invoked 
        /// once and only once, and only if there is an explicit dispose.
        /// </summary>
        protected override void CleanupDisposableObjects()
        {
            
        }

        /// <summary>
        /// Method to clean up native resources.  This method will be invoked
        /// once and only once, regardless of if there is an explicit dispose.
        /// </summary>
        protected override void CleanupUnmanagedResources()
        {
            this.CloseThenKillPowerPoint();
        }
        
        /// <summary>
        /// Method to try to gracefully close powerpoint. The Powerpoint app is known to not close
        /// all the times when you ask it to close. Returns true if app did close, else false.
        /// </summary>
        private bool TryClosePowerPoint(
            CancellationToken token)
        {
            bool result = false;

            ChkArg.IsNotNull(token, nameof(token));
            token.ThrowIfCancellationRequested();
            
            lock (this)
            {
                if (this.ApplicationHandle == null)
                {
                    result = true;
                }
                else
                {
                    try
                    {
                        this.ClosePresentations(null);
                    }
                    catch (Exception ex)
                    {
                        // Log and swallow exception as failures are non-critical to processing
                        Logger.Write("this.ClosePresentations", ex);
                    }

                    // try to quit the application
                    try
                    {
                        this.ApplicationHandle.Quit();

                        result = true;
                    }
                    catch (Exception ex)
                    {
                        // Log and swallow exception as failures are non-critical to processing
                        Logger.Write("this.ApplicationHandle.Quit", ex);
                    }
                }
            }

            return result;
        }

        /// <summary>
        /// Method to attempt a graceful close of powerpoint and then HARD terminate the process(es) if
        /// it did not close gracefully.
        /// </summary>
        private void CloseThenKillPowerPoint()
        {
            if (this.ApplicationHandle != null)
            {
                bool isPptClosed = false;

                CancellationTokenSource tokenSource = new CancellationTokenSource();
                
                // Attempt to kill PPT.
                // We have observed that PPT sometime won't die gracefully. For this reason we
                // have the call to kill PPT off thread and we wait for thread to complete for so long
                // before we hard kill the PPT process.
                // See bug #48247 for more details
                Task task = Task.Run(
                    () =>
                    {
                        isPptClosed = this.TryClosePowerPoint(tokenSource.Token);
                    },
                    tokenSource.Token);
                Task.WhenAny(
                    task,
                    Task.Delay(Settings.Default.PowerpointForceCloseTimeout, tokenSource.Token)).Wait(tokenSource.Token);

                if (!isPptClosed)
                {
                    KillAnyRunningPowerPointInstances();
                }

                if (!task.IsCompleted && !task.IsCanceled && !task.IsFaulted)
                {
                    tokenSource.Cancel();
                }

                lock (this)
                {
                    this.ApplicationHandle = null;
                }
            }
        }

        /// <summary>
        /// Method to close a powerpoint presentation in the active PowerPoint app handle
        /// </summary>
        public void ClosePresentations(
            PptPresentation pptPresentation)
        {
            lock (this)
            {
                // PowerpointPresentation should be the only file open, but just null it out and loop over
                // the entire collection just to be on the safe side
                if (this.ApplicationHandle != null)
                {
                    if (pptPresentation != null)
                    {
                        pptPresentation.Close();
                    }

                    foreach (PptPresentation presentation in this.ApplicationHandle.Presentations)
                    {
                        presentation.Close();
                    }
                }
            }
        }

        /// <summary>
        /// Kill any running PowerPoint instances.
        /// </summary>
        public static void KillAnyRunningPowerPointInstances()
        {
            // ReSharper disable once StringLiteralTypo
            KillProcessByName("powerpnt");

            // also kill Excel, which might come up if the PPT has embedded excel
            // ReSharper disable once StringLiteralTypo
            KillProcessByName("excelcnv");
        }

        /// <summary>
        /// Helper method to hard kill all processes with a matching process name
        /// </summary>
        private static void KillProcessByName(string name)
        {
            foreach (Process process in Process.GetProcessesByName(name))
            {
                try
                {
                    process.Kill();
                }
                catch (Exception ex)
                {
                    // Note: The above with throw if the process died between enumeration and the process.Kill() call
                    Logger.Write("process.Kill", ex);
                }
            }
        }
    }
}