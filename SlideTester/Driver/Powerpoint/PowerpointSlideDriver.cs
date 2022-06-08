using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;

using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using ImageMagick;
using ImageMagick.ImageOptimizers;
using Microsoft.Office.Interop.PowerPoint;

using SlideTester.Common;
using SlideTester.Common.Extensions;
using SlideTester.Common.Log;
using Effect = Microsoft.Office.Interop.PowerPoint.Effect;
using PptApplication = Microsoft.Office.Interop.PowerPoint.Application;
using PptPresentation = Microsoft.Office.Interop.PowerPoint.Presentation;
using PptSlide = Microsoft.Office.Interop.PowerPoint.Slide;
using PanoptoSlide = SlideTester.Driver.Slide;
using Shape = Microsoft.Office.Interop.PowerPoint.Shape;

namespace SlideTester.Driver.Powerpoint
{
    internal class PowerpointSlideDriver : SafeDisposable, ISlideDriver
    {
        #region Members and properties

        private static string PowerPointLockName { get; } = "PowerPointLockName";
        private static string ImageFileExtension { get; } = "jpg";
        
        /// <summary>
        /// RegEx to determine if a path contains a slide image file.
        /// </summary>
        private string ImageRegEx => $@"\d*\.{ImageFileExtension}$";
        private List<FileStream> FileLocks { get; set; } = new List<FileStream>();
        private Mutex PowerpointMutex { get; set; } = null;
        private ScratchSpace Scratch { get; }

        private event EventHandler NonFatalErrorEvent;
        
        #endregion

        [DllImport("kernel32.dll", SetLastError = true)]
        private static extern bool Wow64DisableWow64FsRedirection(ref IntPtr ptr);

        [DllImport("kernel32.dll", SetLastError = true)]
        private static extern bool Wow64RevertWow64FsRedirection(IntPtr ptr);
        
        public PowerpointSlideDriver(
            EventHandler nonFatalErrorEvent = null)
        {
            this.NonFatalErrorEvent = nonFatalErrorEvent;
            this.Scratch = new ScratchSpace();
        }
        
        public Task<List<Slide>> ExtractSlidesAsync(
            string inputFilePath,
            string outputFolderPath,
            CancellationToken token)
        {
            List<Slide> results;

            ChkArg.IsNotNullOrWhiteSpace(inputFilePath, nameof(inputFilePath));
            ChkArg.IsTrue(File.Exists(inputFilePath), nameof(inputFilePath));
            ChkArg.IsNotNullOrWhiteSpace(outputFolderPath, nameof(outputFolderPath));
            ChkArg.IsNotNull(token, nameof(token));
            token.ThrowIfCancellationRequested();

            if (!Directory.Exists(outputFolderPath))
            {
                Directory.CreateDirectory(outputFolderPath);
            }
            
            if (!RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
            {
                ChkExpect.Fail($"OS ({RuntimeInformation.OSDescription}) is not supported for this functionality");
            }

            try
            {
                this.Initialize();
                results = PerformProcessingInternal(inputFilePath, outputFolderPath, token)
                    .OrderBy(s => s.SlideNumber)
                    .ToList();
            }
            finally
            {
                DisposablePptApplication.KillAnyRunningPowerPointInstances();
                this.WaitForGC();
                SafeDisposable.Dispose(this.PowerpointMutex);

                PowerpointMessageFilter.Revoke();

                if (this.FileLocks != null)
                {
                    this.FileLocks.ForEach(fs => fs.Dispose());
                    this.FileLocks.Clear();
                }
            }

            return Task.FromResult(results);
        }

        /// <summary>
        /// Method to clean up managed resources.  This method will be invoked 
        /// once and only once, and only if there is an explicit dispose.
        /// </summary>
        protected override void CleanupDisposableObjects()
        {
            SafeDisposable.Dispose(this.Scratch);
        }

        /// <summary>
        /// Method to clean up native resources.  This method will be invoked
        /// once and only once, regardless of if there is an explicit dispose.
        /// </summary>
        protected override void CleanupUnmanagedResources()
        {
            // Nothing to do
        }
        
        private List<PanoptoSlide> PerformProcessingInternal(
            string inputFilePath,
            string outputFolderPath,
            CancellationToken token)
        {
            List<PanoptoSlide> results = null;
            Exception processingException = null;
            SlideProcessingException.FailureReason? failureReason = null;
            
            ChkArg.IsNotNull(token, nameof(token));
            token.ThrowIfCancellationRequested();
            
            if (this.IsPasswordProtected(inputFilePath))
            {
                throw new SlideProcessingException(
                    inputFilePath,
                    SlideProcessingException.FailureReason.PasswordProtected);
            }

            try
            {
                // Office runs as a STA application it only has one thread dedicated to all of its processing, this means
                // that registering a IOleMessageFilter with it requires that the thread we are registering in also be
                // a STA thread. If not the registration will silently fail and the retry mechanism will not work.
                //
                // This is why this method now runs in its own thread and is set to be a STA thread.
                Thread staOfficeThread = new Thread(
                    delegate()
                    {
                        (results, processingException) = PerformProcessing_STAWorker(
                            inputFilePath,
                            outputFolderPath,
                            token);
                    });
                staOfficeThread.SetApartmentState(ApartmentState.STA);

                // Start processing the powerpoint slides if we do not return in a set amount of time abort the thread
                // and mark the job as failed.
                staOfficeThread.Start();
                if (!staOfficeThread.Join(Settings.Default.ProcessingTimeout))
                {
                    staOfficeThread.Abort();
                    failureReason = SlideProcessingException.FailureReason.TimedOut;
                }
                else if (processingException != null)
                {
                    failureReason = SlideProcessingException.FailureReason.FailureUnknown;
                }
            }
            catch (Exception ex)
            {
                throw new SlideProcessingException(
                    inputFilePath,
                    SlideProcessingException.FailureReason.ImageExtractionFailure,
                    ex);
            }

            if (failureReason.HasValue)
            {
                throw new SlideProcessingException(inputFilePath, failureReason.Value, processingException);
            }
            
            return results ?? new List<PanoptoSlide>();
        }
        
        private void Initialize()
        {
            // if we are running as "SYSTEM", hold a handle in systemProfile\Desktop
            // so that ppt doesn't fall over. This ends up a bit complex due to all the architecture combos
            if (System.Security.Principal.WindowsIdentity.GetCurrent().IsSystem)
            {
                ChkExpect.IsTrue(Environment.Is64BitOperatingSystem, "Environment.Is64BitOperatingSystem");
                ChkExpect.IsTrue(Environment.Is64BitProcess, "Environment.Is64BitProcess");

                IntPtr wow64State = new IntPtr();
                Wow64DisableWow64FsRedirection(ref wow64State);

                try
                {
                    this.FileLocks.Add(
                        OpenFileLockInSystemProfileDesktop(
                            Environment.GetFolderPath(Environment.SpecialFolder.System)));
                    this.FileLocks.Add(
                        OpenFileLockInSystemProfileDesktop(
                            Environment.GetFolderPath(Environment.SpecialFolder.SystemX86)));
                }
                finally
                {
                    Wow64RevertWow64FsRedirection(wow64State);
                }
            }
        }

        private void OpenPowerPoint(
            out DisposablePptApplication powerPointApplication)
        {
            // create a PowerPoint application
            powerPointApplication = new DisposablePptApplication(new PptApplication());

            // This is required to disable macros from running, which are a serious security hole
            powerPointApplication.ApplicationHandle.AutomationSecurity =
                Microsoft.Office.Core.MsoAutomationSecurity.msoAutomationSecurityForceDisable;

            powerPointApplication.ApplicationHandle.Activate();
        }

        /// <summary>
        /// Helper to open a file in the system desktop folder
        /// Powerpoint is a user mode application and weirdly has a dependency on that folder existing
        /// </summary>
        private FileStream OpenFileLockInSystemProfileDesktop(string baseSystemPath)
        {
            // ReSharper disable once StringLiteralTypo
            string systemDesktopDir = Path.Combine(baseSystemPath, "config\\systemprofile\\Desktop");

            if (!Directory.Exists(systemDesktopDir))
            {
                Directory.CreateDirectory(systemDesktopDir);
            }

            // open up the file lock using a temporary file handle and a temporary name.
            return File.Open(Path.Combine(
                    systemDesktopDir,
                    Guid.NewGuid().ToString()),
                FileMode.Create,
                FileAccess.ReadWrite,
                FileShare.None);
        }

        /// <summary>
        /// Detects if the associated this.PowerpointFilePath is password protected
        /// See https://stackoverflow.com/questions/5778074/detecting-a-password-protected-document for details
        /// </summary>
        /// <returns>true is guaranteed to be password protected, false otherwise</returns>
        private bool IsPasswordProtected(
            string powerpointFilePath)
        {
            using Stream fileStream = File.OpenRead(powerpointFilePath);
            
            // minimum file size for supported office files is 4k
            if (fileStream.Length < 4096)
            {
                return false;
            }

            // read file header
            fileStream.Seek(0, SeekOrigin.Begin);
            var compObjHeader = new byte[0x20];
            if (!this.ReadEntireBufferFromStream(fileStream, compObjHeader))
            {
                return false;
            }

            // check if we have plain zip file
            if (compObjHeader[0] == 'P' && compObjHeader[1] == 'K')
            {
                // this is a plain OpenXml document (not encrypted)
                return false;
            }

            // check compound object magic bytes (used in legacy PPT files)
            if (compObjHeader[0] != 0xD0 || compObjHeader[1] != 0xCF)
            {
                // unknown document format
                return false;
            }

            int sectionSizePower = compObjHeader[0x1E];
            if (sectionSizePower < 8 || sectionSizePower > 16)
            {
                // invalid section size
                return false;
            }

            // scan 32K at a time, which is sufficient to find the header/footer we're looking for
            const int defaultScanLength = 32768;
            long scanLength = Math.Min(defaultScanLength, fileStream.Length);

            // read header part for scan
            fileStream.Seek(0, SeekOrigin.Begin);
            var header = new byte[scanLength];
            if (!ReadEntireBufferFromStream(fileStream, header))
            {
                return false;
            }

            // check if we detected password protection
            if (ScanForPasswordProtection(header))
            {
                return true;
            }

            // if not, try to scan footer as well

            // read footer part for scan
            fileStream.Seek(-scanLength, SeekOrigin.End);
            var footer = new byte[scanLength];
            if (!ReadEntireBufferFromStream(fileStream, footer))
            {
                return false;
            }

            // return the final result
            return ScanForPasswordProtection(footer);
        }
        
        /// <summary>
        /// Scans for any known markers that the file is password protected. Flexible to support multiple powerpoint formats
        /// </summary>
        /// <param name="stream">Underlying stream - may be seeked and modified during execution</param>
        /// <param name="buffer">Input data buffer</param>
        /// <param name="sectionSize">How large the section to search is</param>
        /// <returns>True if password protected, false otherwise</returns>
        private bool ScanForPasswordProtection(byte[] buffer)
        {
            const string afterNamePadding = "\0\0\0\0\0\0\0\0\0\0\0\0\0\0\0\0";

            try
            {
                string bufferString = Encoding.ASCII.GetString(buffer, 0, buffer.Length);

                // try to detect password protection used in new OpenXml documents
                // by searching for "EncryptedPackage" or "EncryptedSummary" streams
                const string encryptedPackageName = "E\0n\0c\0r\0y\0p\0t\0e\0d\0P\0a\0c\0k\0a\0g\0e" + afterNamePadding;
                const string encryptedSummaryName = "E\0n\0c\0r\0y\0p\0t\0e\0d\0S\0u\0m\0m\0a\0r\0y" + afterNamePadding;
                if (bufferString.Contains(encryptedPackageName)
                    || bufferString.Contains(encryptedSummaryName))
                {
                    return true;
                }

                // we're currently unable to detect password protection for legacy Powerpoint documents
                // so take no action here (the stackoverflow approach doesn't actually work properly)
                // tfs-79884
            }
            catch (ArgumentException)
            {
                // BitConverter exceptions may be related to document format problems
                // so we just treat them as "password not detected" result
                return false;
            }

            return false;
        }
        
        /// <summary>
        /// Reads a given buffer from a stream, retrying to ensure the requested buffer is entirely read
        /// </summary>
        /// <returns>False if the end of stream was hit and we couldn't read the entire buffer</returns>
        private bool ReadEntireBufferFromStream(Stream stream, byte[] buffer)
        {
            int bytesRemaining = buffer.Length;
            while (bytesRemaining > 0)
            {
                int bytesRead = stream.Read(buffer, 0, bytesRemaining);
                if (bytesRead == 0)
                {
                    break;
                }

                bytesRemaining -= bytesRead;
            }

            return bytesRemaining == 0;
        }

        private (List<PanoptoSlide>, Exception ex) PerformProcessing_STAWorker(
            string inputFilePath,
            string outputFolderPath,
            CancellationToken token)
        {
            List<PanoptoSlide> results = null;
            Exception exception = null;

            ChkArg.IsNotNull(token, nameof(token));
            token.ThrowIfCancellationRequested();

            try
            {
                using DisposablePptApplication pptApplication = this.ConnectToPowerpoint(inputFilePath, token);

                bool succeeded;
                try
                {
                    results = this.ExtractSlidesFromPowerpoint(
                        pptApplication,
                        inputFilePath,
                        outputFolderPath,
                        shouldScrubPpt: true,
                        token);
                    succeeded = true;
                }
                catch
                {
                    // Catch and swallow
                    // If we didn't succeed with scrubbing our PPT, we will retry below w/o scrubbing
                    succeeded = false;
                }

                if (!succeeded)
                {
                    results = this.ExtractSlidesFromPowerpoint(
                        pptApplication,
                        inputFilePath,
                        outputFolderPath,
                        shouldScrubPpt: false,
                        token);
                }
            }
            catch (Exception ex)
            {
                exception = ex;
            }

            return (results ?? new List<PanoptoSlide>(), exception);
        }
        
        private List<PanoptoSlide> ExtractSlidesFromPowerpoint(
            DisposablePptApplication pptApplication,
            string inputFilePath,
            string outputFolderPath,
            bool shouldScrubPpt,
            CancellationToken token)
        {
            PptPresentation pptPresentation = null;
            List<PanoptoSlide> results;

            ChkArg.IsNotNull(pptApplication, nameof(pptApplication));
            ChkArg.IsNotNull(token, nameof(token));
            token.ThrowIfCancellationRequested();
            
            try
            {
                pptPresentation = this.OpenFile(pptApplication, inputFilePath, shouldScrubPpt, token);
                
                Size unscaledSize = new Size(
                    (int)(pptPresentation.PageSetup.SlideWidth + 0.5f),
                    (int)(pptPresentation.PageSetup.SlideHeight + 0.5f));
                uint maxPixels = ResolutionHelper.MaxPixelCount(unscaledSize);

                Size desiredSize = unscaledSize.ComputeScaledDimensions(maxPixels);

                // Some powerpoint versions (at least 2k12 and not 2k16) have a bug where it
                // will always round up the width to the nearest multiple of 8
                // and it will always round height up to the nearest multiple of 4 when determining the output
                // canvas size. The images exported will always match the resolutions we pass though, so we may get
                // some black edges around our images if our supplied width and height are not the correct multiples.
                // The logic above to get scaled resolution should make it so this doesn't happen for all the most common aspect ratio
                // but there are zillions of ARs so we just add the buffer now for the odd-ball ones.
                desiredSize.Width += desiredSize.Width % 8;
                desiredSize.Height += desiredSize.Height % 4;                
                
                results = this.ExtractSlidesContent(
                    pptPresentation,
                    inputFilePath, 
                    outputFolderPath,
                    desiredSize,
                    token).Result;
            }
            finally
            {
                // regardless of success or failure, attempt to close the PPT
                // if the object is null, this is still safe.
                pptApplication.ClosePresentations(pptPresentation);
            }

            return results;
        }
        
        private DisposablePptApplication ConnectToPowerpoint(
            string inputFilePath,
            CancellationToken token)
        {
            DisposablePptApplication pptApplication = null;

            ChkArg.IsNotNull(token, nameof(token));
            token.ThrowIfCancellationRequested();
            
            DisposablePptApplication.KillAnyRunningPowerPointInstances();
            this.WaitForGC();

            this.PowerpointMutex = new Mutex(initiallyOwned:false, PowerPointLockName);
            bool hasLock = this.PowerpointMutex.WaitOne(Settings.Default.PowerpointAcquireLockTimeout);

            if (!hasLock)
            {
                Logger.Write($"Packager_ProcessPowerPointTarget_CompleteInitFailure");
                throw new SlideProcessingException(
                    inputFilePath,
                    SlideProcessingException.FailureReason.LockNotAcquired,
                    "Failed to get PPT Lock.");
            }

            // PPT can throw threading errors, so register a message filter that always retries.
            // Make sure to capture the return HResult of the registration method, if it doesn't return 0 it failed
            // and it won't properly retry.
            uint comRegisterMessageFilterHResult = (uint)PowerpointMessageFilter.Register();
            if (comRegisterMessageFilterHResult == 0)
            {
                bool success = false;
                DisposablePptApplication.KillAnyRunningPowerPointInstances();
                List<Exception> exceptions = new List<Exception>();

                int iRetryCount = Settings.Default.ComRetryCount;

                while (!success && (iRetryCount > 0))
                {
                    try
                    {
                        this.OpenPowerPoint(out pptApplication);
                        success = true;
                    }
                    catch (Exception e)
                    {
                        exceptions.Add(e);

                        // drop our retry counter, sleep for a bit, and give it another go.
                        iRetryCount--;
                        Thread.Sleep(Settings.Default.ComRetryDelay);
                    }
                }

                if (!success)
                {
                    throw new SlideProcessingException(
                        inputFilePath,
                        SlideProcessingException.FailureReason.PowerpointStartFailure,
                        "Failed to open powerpoint",
                        new AggregateException(exceptions));
                }
            }
            else
            {
                Logger.Write($"PPT_IOleMessageFilterRegistration_Failed {comRegisterMessageFilterHResult:X}");
                throw new SlideProcessingException(
                    inputFilePath,
                    SlideProcessingException.FailureReason.PowerpointComRegisterFailure,
                    $"Failure to register OleMessageFilter (code={comRegisterMessageFilterHResult:X})");
            }

            return pptApplication;
        }
        
        /// <summary>
        /// Wait for garbage collection to complete.  This is necessary to allow graceful shutdown of Powerpoint
        /// after we've asked it to shut down.  Under the hood, there are a few interface pointers that need to
        /// be released to complete graceful shutdown; these are handled by the finalizer thread.
        ///
        /// CAUTION: Do not, directly, or indirectly, call or wait on a thread calling GC.WaitForPendingFinalizers()
        /// from the STA thread. If Powerpoint has become non-responsive, we end up blocking the STA message pump
        /// forever waiting for RPC calls to Powerpoint that will never return.
        ///
        /// NOTE: If we fail to wait for GC to complete, we will kill all running Powerpoint instances.
        /// </summary>
        private void WaitForGC()
        {
            Thread cleanupThread = new Thread(() =>
            {
                GC.Collect();
                GC.WaitForPendingFinalizers();
            });

            cleanupThread.Start();

            if (!cleanupThread.Join(Settings.Default.PowerpointForceCloseTimeout))
            {
                cleanupThread.Abort();

                Logger.Write($"PPT_WaitForGC_Timeout {Settings.Default.PowerpointForceCloseTimeout}");

                DisposablePptApplication.KillAnyRunningPowerPointInstances();
            }
        }
        
        /// <summary>
        /// Attempt to open the specified PPT file.
        /// </summary>
        /// <param name="strFileName">File to open.</param>
        /// <param name="tryScrubFile">True if we should to scrub the file by opening it and
        /// saving a copy.  False otherwise.</param>
        /// <returns>True if the file was opened successfully. False if there was an error.</returns>
        /// <remarks>Call this function first with tryScrubFile set to true.  If that fails,
        /// call it again with tryScrubFile set to false.</remarks>
        private PptPresentation OpenFile(
            DisposablePptApplication powerPointApplication,
            string inputFilePath,
            bool shouldScrubDocument,
            CancellationToken token)
        {
            PptPresentation pptPresentation = null;

            ChkArg.IsNotNull(powerPointApplication, nameof(powerPointApplication));
            ChkArg.IsNotNull(token, nameof(token));
            token.ThrowIfCancellationRequested();
            
            try
            {
                string pptFileToProcess = inputFilePath; 
                
                if (shouldScrubDocument)
                {
                    if (this.TryScrubPPTFile(powerPointApplication, inputFilePath, token, out string scrubbedDocument))
                    {
                        pptFileToProcess = scrubbedDocument;
                    }
                }

                // open from the temporary file.
                pptPresentation = powerPointApplication.ApplicationHandle.Presentations.Open(
                    pptFileToProcess,
                    Microsoft.Office.Core.MsoTriState.msoTrue,
                    Microsoft.Office.Core.MsoTriState.msoFalse,
                    Microsoft.Office.Core.MsoTriState.msoFalse);
            }
            catch (Exception ex)
            {
                Logger.Write($"PPT_TryOpenFile_UnhandledException {inputFilePath}, {shouldScrubDocument}", ex);
            }

            return pptPresentation;
        }
                
        /// <summary>
        /// Open a copy of the specified PPT file and save it with a different file name.
        ///
        /// On success:
        /// returns true
        /// outputFileName is set to a temporary file name.
        ///
        /// On failure:
        /// returns false
        /// outputFileName is set to an empty string (not null).
        /// </summary>
        /// <param name="strFileName">PPT file to copy.</param>
        /// <remarks>The purpose of this function is to try to get PPT to scrub the file
        /// using the Presentation.SaveAs() function.</remarks>
        private bool TryScrubPPTFile(
            DisposablePptApplication powerpointApplication,
            String inputFileName,
            CancellationToken token,
            out string outputFileName)
        {
            ChkArg.IsNotNull(token, nameof(token));
            token.ThrowIfCancellationRequested();
            
            // important to prepend filename, otherwise powerpoint tries to be too smart for us during save as
            // and may change the file extension between .ppt <=> .pptx
            string dirName = Path.GetDirectoryName(inputFileName);
            outputFileName = null;
            string tempFileName = Path.Combine(
                dirName ?? throw new ArgumentException("inputFileName ({inputFileName}) contains no folder prefix"), 
                "temp_" + Path.GetFileName(inputFileName));
            PptPresentation pTempPres = null;
            bool succeeded = false;
            
            try
            {
                // open the supplied filename and save out to temp location.
                pTempPres = powerpointApplication.ApplicationHandle.Presentations.Open(
                    inputFileName,
                    Microsoft.Office.Core.MsoTriState.msoTrue,
                    Microsoft.Office.Core.MsoTriState.msoFalse,
                    Microsoft.Office.Core.MsoTriState.msoFalse);

                pTempPres.SaveAs(tempFileName, PpSaveAsFileType.ppSaveAsPresentation, Microsoft.Office.Core.MsoTriState.msoTrue);

                // always re-read the actual new filename back out.
                outputFileName = pTempPres.FullName;

                succeeded = true;
            }
            catch (Exception ex)
            {
                // we want to log and return a failure. Note that this method is not critical if it fails as we
                // know that scrubbing fails some % of the time.
                Logger.Write($"PPT_CopyFile_CreateScrubbedPPTFailure {inputFileName}", ex);
                outputFileName = string.Empty;
            }
            finally
            {
                try
                {
                    if (pTempPres != null)
                    {
                        // close the temp file.
                        pTempPres.Close();
                    }
                }
                catch (Exception ex)
                {
                    // We want to log and swallow if we failed to close the presentation as this is "interesting" but not critical
                    Logger.Write($"PPT_CopyFile_CloseTempPPTFailure", ex);
                }
            }

            return succeeded;
        }
        
        /// <summary>
        /// Saves the currently loaded PPT file as a series of image files in the target directory.
        /// Returns a dictionary of paths to the generated slide images by their slideNumbers
        /// </summary>
        private Dictionary<int, string> ExtractSlideImages(
            PptPresentation pptPresentation,
            string inputFilePath,
            string outputFolderPath,
            Size desiredSize,
            CancellationToken token)
        {
            var results = new Dictionary<int, string>();

            ChkArg.IsNotNullOrWhiteSpace(outputFolderPath, nameof(outputFolderPath));
            ChkArg.IsTrue(Directory.Exists(outputFolderPath), nameof(outputFolderPath));
            ChkArg.IsNotNull(pptPresentation, nameof(pptPresentation));
            ChkArg.IsNotNull(token, nameof(token));
            token.ThrowIfCancellationRequested();
            
            try
            {
                ChkExpect.IsNotNull(pptPresentation, "There is no open PowerPoint presentation.");

                // Export will throw if there are zero slides.
                if (pptPresentation.Slides.Count > 0)
                {
                    string tempOutputPath = this.Scratch.CreateUniqueFolder();
                    
                    pptPresentation.Export(
                        tempOutputPath,
                        ImageFileExtension,
                        desiredSize.Width,
                        desiredSize.Height);

                    // ReSharper disable once CommentTypo
                    // The french version of ppt outputs files as "Diapositive<#>.jpg",
                    // while other parts of the pipeline expect "slide<#>.jpg".
                    // To fix, normalize all files of format "<word><#>.jpg" to "slide<#>.jpg"
                    foreach (string imagePath in Directory.EnumerateFiles(tempOutputPath))
                    {
                        // search for the "<#>.jpg" stem
                        Match match = Regex.Match(imagePath, this.ImageRegEx, RegexOptions.IgnoreCase);
                        if (match.Success)
                        {
                            // put on the Slide prefix expected by other parts of the pipeline
                            string normalizedFile = Path.Combine(
                                outputFolderPath, 
                                $"slide{match.Value}").ToLowerInvariant();

                            if (!imagePath.Equals(normalizedFile, StringComparison.InvariantCulture))
                            {
                                // move the file if it is different from the original
                                if (File.Exists(normalizedFile))
                                {
                                    // Console.WriteLine($"Warning! Overwriting file: {normalizedFile}");
                                    File.Delete(normalizedFile);
                                }
                                
                                File.Move(imagePath, normalizedFile);
                            }

                            string slideNumString = match.Value.Split('.').First();
                            int slideNum = int.Parse(slideNumString);
                            
                            ChkExpect.IsTrue(
                                !results.ContainsKey(slideNum), 
                                $"Image for slide (number:{slideNum}) has already been extracted");
                            
                            results[slideNum] = normalizedFile;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                throw new SlideProcessingException(
                    inputFilePath,
                    SlideProcessingException.FailureReason.ImageExtractionFailure,
                    ex);
            }

            return results;
        }
        
        /// <summary>
        /// Extracts slide text, animations and notes.
        /// </summary>
        private async Task<List<PanoptoSlide>> ExtractSlidesContent(
            PptPresentation pptPresentation,
            string inputFilePath,
            string outputFilePath,
            Size desiredSize,
            CancellationToken token)
        {
            var results = new List<PanoptoSlide>(); 

            ChkArg.IsNotNull(pptPresentation, nameof(pptPresentation));
            ChkArg.IsNotNull(token, nameof(token));
            token.ThrowIfCancellationRequested();

            try
            {
                foreach (PptSlide slide in pptPresentation.Slides)
                {
                    List<string> title;
                    List<string> subtitle;
                    List<string> textShapes = null;
                    List<string> animationsText = null;
                    List<string> notes = null;

                    string slideImagePath;

                    if (!Directory.Exists(outputFilePath))
                    {
                        Directory.CreateDirectory(outputFilePath);
                    }
                    
                    try
                    {
                        slideImagePath = Path.Combine(
                            outputFilePath,
                            $"{slide.SlideNumber}-{Guid.NewGuid()}.png").ToLowerInvariant();
                        await this.ExportSlideImage(slideImagePath, desiredSize, slide, token).ConfigureAwait(false);
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Failed to export image for slide #{slide.SlideNumber}. Exception:\n{ex}");
                        slideImagePath = string.Empty;
                    }
                    
                    // get the slide title
                    try
                    {
                        title =
                            (from Shape shape in slide.Shapes
                             where
                                shape.Name.Contains("Title") // a hack to pick up special title shapes
                                && !shape.Name.Contains("Subtitle") // skip subtitles
                                && shape.HasTextFrame == Microsoft.Office.Core.MsoTriState.msoTrue
                                && (shape.TextFrame.TextRange != null)
                             select shape.TextFrame.TextRange.Text).ToList();
                    }
                    catch (Exception ex)
                    {
                        // Slide processing requires a title. Log the exception and continue to next slide
                        this.NonFatalErrorEvent?.Invoke(
                            this, 
                            new NonFatalErrorEventArgs(
                                "Failed to extract title. Skipping further processing of slide",
                                slide.SlideNumber,
                                 ex));
                        continue;
                    }

                    // get the slide title
                    try
                    {
                        subtitle =
                            (from Shape shape in slide.Shapes
                                where
                                    shape.Name.Contains("Subtitle") // a hack to pick up special title shapes
                                    && shape.HasTextFrame == Microsoft.Office.Core.MsoTriState.msoTrue
                                    && (shape.TextFrame.TextRange != null)
                                select shape.TextFrame.TextRange.Text).ToList();
                    }
                    catch (Exception ex)
                    {
                        // Slide processing requires a title. Log the exception and continue to next slide
                        this.NonFatalErrorEvent?.Invoke(
                            this, 
                            new NonFatalErrorEventArgs(
                                "Failed to extract title. Skipping further processing of slide",
                                slide.SlideNumber,
                                ex));
                        continue;
                    }
                    
                    // get slide text shapes
                    try
                    {
                        textShapes =
                            (from Shape shape in slide.Shapes
                             where
                                 // ignore slide numberings
                                 !shape.Name.Contains("Slide Number")
                                 // Skip titles and subtitles as they are included above
                                 && !shape.Name.Contains("Title")
                                 && !shape.Name.Contains("Subtitle")
                                 // and skip animations, we'll get those next
                                 && (shape.AnimationSettings.Animate == Microsoft.Office.Core.MsoTriState.msoFalse)
                                 // only if the slide has text
                                 && (shape.HasTextFrame == Microsoft.Office.Core.MsoTriState.msoTrue)
                                 && (shape.TextFrame.TextRange != null)
                             orderby shape.ZOrderPosition ascending
                             select shape.TextFrame.TextRange.Text).ToList();
                    }
                    catch (Exception ex)
                    {
                        // Failed to grab static text from slide, this is a non-critical error for processing this slide.
                        this.NonFatalErrorEvent?.Invoke(
                            this, 
                            new NonFatalErrorEventArgs(
                                "Failed to extract static text from slide. Continuing processing of slide",
                                slide.SlideNumber,
                                ex));
                    }

                    // step-by-step animation text
                    try
                    {
                        animationsText =
                            (from Effect effect in slide.TimeLine.MainSequence
                             orderby effect.Index ascending
                             select effect.DisplayName).ToList();
                    }
                    catch (Exception ex)
                    {
                        // Failed to grab text from animation slide, this is a non-critical error for processing this slide.
                        this.NonFatalErrorEvent?.Invoke(
                            this, 
                            new NonFatalErrorEventArgs(
                                "Failed to extract animated text from slide. Continuing processing of slide",
                                slide.SlideNumber,
                                ex));
                    }

                    // presenter notes
                    try
                    {
                        notes = (from PptSlide notesPage in slide.NotesPage
                         from Shape shape in notesPage.Shapes
                         where
                             shape.Name.Contains("Notes") // a hack to check to verify that this is a notes shape
                             && shape.HasTextFrame == Microsoft.Office.Core.MsoTriState.msoTrue
                             && (shape.TextFrame.TextRange != null)
                         select shape.TextFrame.TextRange.Text).ToList();
                    }
                    catch (Exception ex)
                    {
                        // Failed to grab text from animation  slide, this is a non-critical error for processing this slide.
                        this.NonFatalErrorEvent?.Invoke(
                            this, 
                            new NonFatalErrorEventArgs(
                                "Failed to extract presenter notes from slide. Continuing processing of slide",
                                slide.SlideNumber,
                                ex));
                    }

                    results.Add(new Slide(
                        slide.SlideNumber,
                        slideImagePath,
                        //imagesBySlideId[slide.SlideNumber],
                        title,
                        subtitle,
                        headers: Enumerable.Empty<string>(),
                        footers: Enumerable.Empty<string>(),
                        bodyText: textShapes,
                        animationText: animationsText,
                        presenterNotes: notes,
                        otherText: Enumerable.Empty<string>()));
                }
            }
            catch (Exception ex)
            {
                throw new SlideProcessingException(
                    inputFilePath,
                    SlideProcessingException.FailureReason.TextExtractionFailure,
                    ex);
            }

            return results;
        }

        private async Task ExportSlideImage(
            string outputFilePath,
            Size desiredSize,
            PptSlide slide,
            CancellationToken token)
        {
            string tempFile = this.Scratch.UniqueFilePath("bmp");
            slide.Export(tempFile, "bmp");

            using MagickImage slideImage = new MagickImage(tempFile);
            if (slideImage.Width != desiredSize.Width || slideImage.Height != desiredSize.Height)
            {
                slideImage.Resize(desiredSize.Width, desiredSize.Height);
            }
            
            await using Stream pngStream = File.Create(outputFilePath);
            PngOptimizer optimizer = new PngOptimizer();
            await slideImage.WriteAsync(pngStream, MagickFormat.Png, token).ConfigureAwait(false);
            
            File.Delete(tempFile);
        }
    }
}