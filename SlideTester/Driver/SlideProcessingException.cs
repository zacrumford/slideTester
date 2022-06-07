using System;

namespace SlideTester.Driver
{
    /// <summary>
    /// Represents a known reason to fail slide processing.
    /// </summary>
    [Serializable]
    public sealed class SlideProcessingException : Exception
    {
        public string PowerpointFile { get; }
        
        public enum FailureReason
        {
            PasswordProtected,
            TimedOut,
            ImageExtractionFailure,
            TextExtractionFailure,
            LockNotAcquired,
            PowerpointStartFailure,
            PowerpointComRegisterFailure,
            FailureUnknown,
        }
        public FailureReason Reason { get; }
        
        public SlideProcessingException(
            string powerpointFile,
            FailureReason reason,
            Exception inner = null)
            : base($"Exception while processing slide deck: {powerpointFile}. Reason: {reason}", inner)
        {
            this.Reason = reason;
            this.PowerpointFile = powerpointFile;
        }
        
        public SlideProcessingException(
            string powerpointFile,
            FailureReason reason,
            string message,
            Exception inner = null)
            : base($"Exception while processing slide deck: {powerpointFile}. Reason: {reason}. Message: {message}", inner)
        {
            this.Reason = reason;
            this.PowerpointFile = powerpointFile;
        }
    }
}