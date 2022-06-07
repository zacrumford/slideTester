using System;

namespace SlideTester.Driver.Powerpoint
{
    /// <summary>
    /// Simple class to pass data about a non-critical failure during a slide processing operation 
    /// </summary>
    public class NonFatalErrorEventArgs : EventArgs
    {
        /// <summary>
        /// Failure message
        /// </summary>
        public string Message { get; }
        
        /// <summary>
        /// Slide number which encountered non-critical failure while processing
        /// </summary>
        public int SlideNumber { get; }
        
        /// <summary>
        /// Exception associated with non-critical failure. May be null.
        /// </summary>
        public Exception Exception { get; }
        
        /// <summary>
        /// Default values ctor
        /// </summary>
        public NonFatalErrorEventArgs(
            string message,
            int slideNumber,
            Exception ex)
        {
            this.Message = message;
            this.SlideNumber = slideNumber;
            this.Exception = ex;
        }
    }
}