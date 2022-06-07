using System;
using System.Threading;
using System.Threading.Tasks;

using CommandLine;

using SlideTester.App.CommandLine;
using SlideTester.Common;
using SlideTester.Common.Log;

namespace SlideTester.App;
public static class SlideTesterProgram
{
    private static CancellationTokenSource CancellationSource { get; } = new CancellationTokenSource();
    private static TimeSpan MaxShutdownDelay { get; } = TimeSpan.FromSeconds(5);
    private static ManualResetEvent MainThreadCompletedEvent { get; } = new ManualResetEvent(initialState: false);

    public static int Main(string[] args)
    {
        int returnCode = 0;
        try
        {
            returnCode = PerformProcessingWrapper(args);
        }
        catch (Exception ex)    
        {
            Logger.Write("BackendTaskWorker_UnhandledTopLevelException", ex);
            returnCode = -3;
            SlideTesterProgram.CancellationSource.Cancel();
        }
        finally
        {
            // The main method is now exiting.  This does not necessarily mean that the process has exited, as we
            // may have other threads running.
            Logger.Write($"Common_ProgramEnding {returnCode}");
            
            // Notify the other threads that this method is completed.
            SlideTesterProgram.MainThreadCompletedEvent.Set();
            
            // Ensure that all extraneous threads are stopped.  They should all be wired into this.
            if (!SlideTesterProgram.CancellationSource.IsCancellationRequested)
            {
                SlideTesterProgram.CancellationSource.Cancel();
            }
        }

        return returnCode;
    }
    
    private static int PerformProcessingWrapper(string [] args)
    {
        try
        {
            // This wires up our graceful shutdown.  It should be the first part of the main method.
            AppDomain.CurrentDomain.ProcessExit += SlideTesterProgram.OnProcessExit;
            Console.CancelKeyPress += SlideTesterProgram.OnCancelKeyPress;
            AppDomain.CurrentDomain.UnhandledException += SlideTesterProgram.ProcessUnhandledException;

            if (SlideTesterProgram.CancellationSource.IsCancellationRequested)
            {
                return -1;
            }

            IVerbProcessor processor = null;

            Parser.Default.ParseArguments<CsvOptions, UriOptions>(args)
                .WithParsed<CsvOptions>(options => { processor = new CsvProcessor(options); })
                .WithParsed<UriOptions>(options => { processor = new UriProcessor(options); });

            if (processor != null)
            {
                // Kick off our work on a different thread so that we can wait on key press and completion
                Task doWorkTask = Task.Run(
                    async () => await processor.DoWork(CancellationSource.Token).ConfigureAwait((false)), 
                    CancellationSource.Token);
                
                WaitHandle.WaitAny(new []
                {
                    ((IAsyncResult)doWorkTask).AsyncWaitHandle,
                    CancellationSource.Token.WaitHandle
                });

                if (doWorkTask.IsFaulted)
                {
                    Console.WriteLine($"Processing failed.\n{doWorkTask.Exception?.ToString() ?? string.Empty }");
                }
            }
            else
            {
                ChkExpect.Fail("Invalid args");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Top level exception:\n{ex}");
            return -2;
        }

        return 0;
    }
    
    /// <summary>
    /// Handler to catch an log unhandled exceptions on worker threads
    /// </summary>
    static void ProcessUnhandledException(object sender, UnhandledExceptionEventArgs unhandledExceptionEventArgs)
    {
        Exception ex = null;
        if (unhandledExceptionEventArgs != null)
        {
            ex = unhandledExceptionEventArgs.ExceptionObject as Exception;
        }

        // This is an unhandled top level exception on this thread
        // log it and always echo this one to console too as we want it as discoverable as possible
        Logger.Write("BackendTaskWorker_UnhandledWorkerThreadException", ex);
        
        // Something is obviously wrong, so shut everything down.
        SlideTesterProgram.CancellationSource.Cancel();
    }

    /// <summary>
    /// On Linux is the handler for SIGINT.
    /// On Windows this is the handler for Ctrl+C or Ctrl+Break. 
    /// </summary>
    private static void OnCancelKeyPress(object sender, ConsoleCancelEventArgs args)
    {
        // Log the fact that we are responding.
        Logger.Write("BackendTaskWorker_CaughtSigint");
        
        Console.WriteLine("This application has caught an interrupt signal and is attempting graceful shutdown.");
        Console.WriteLine($"If necessary, forceful shutdown will occur within {MaxShutdownDelay.TotalSeconds:0.0} seconds.");
        
        // This indicates that the main process should resume, so it has a chance to wind down gracefully.
        args.Cancel = true;
        
        // This should cause all worker threads to halt with exceptions bubbling up to the main thread, which
        // will ultimately cause the main thread to stop (after it handles and logs the exceptions).
        SlideTesterProgram.CancellationSource.Cancel();
    }
    
    /// <summary>
    /// This method is only called once, near the end of an application's lifecycle.
    ///
    /// On Linux this method is called when the process has received a SIGTERM message, which means we need to trigger
    /// a graceful shutdown to bring the process to an end.  This is also how Docker tries to shut down the process
    /// gracefully, for example when `docker stop` is called.
    /// 
    /// Otherwise, if the process did not receive SIGTERM, then this method is called immediately before the process
    /// ends.
    /// </summary>
    private static void OnProcessExit(object sender, EventArgs e)
    {
        Logger.Write("Common_StartingGracefulShutdown");
        
        // If we are responding to a SIGTERM, then we need to shut down all the worker threads, which should all
        // be wired into this.  Otherwise, if we are shutting down normally, then the worker threads have already
        // stopped and this call has no effect.
        if (!CancellationSource.IsCancellationRequested)
        {
            // We generally should be cancelled already on graceful shutdown, if we have not cancelled
            // the do so now and log
            SlideTesterProgram.CancellationSource.Cancel();
        }

        // Check to make sure that the main method has finished.
        if (!SlideTesterProgram.MainThreadCompletedEvent.WaitOne(TimeSpan.Zero))
        {
            // This means that main method has not finished, so we are likely handling a SIGTERM.
            
            // Log the fact that this is happening.
            Logger.Write("BackendTaskWorker_PauseDuringGracefulShutdown");
            
            // Wait a bit.  Hopefully the main method will finish.
            SlideTesterProgram.MainThreadCompletedEvent.WaitOne(MaxShutdownDelay);
        }

        Logger.Write("Common_CompletedGracefulShutdown");
    }
}
