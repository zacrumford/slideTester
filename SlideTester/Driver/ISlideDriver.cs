using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;

using Panopto.Common.Media.Slides;

namespace SlideTester.Driver
{
    /// <summary>
    /// Interface representing an object which can extract data from a slide deck file.
    /// </summary>
    public interface ISlideDriver
    {
        /// <summary>
        /// Method which accepts a slide deck file path and extracts all text information
        /// and slide images from the slide deck. Images will be written to the output path. All slide
        /// metadata, including per-slide image file paths, will be written to output.
        /// </summary>
        /// <param name="inputFilePath">Path to the slide deck which will be processed</param>
        /// <param name="outputFolderPath">Path to which generated slide images will be written</param>
        /// <param name="token">Cancellation token to use</param>
        /// <returns>A list of Slide objects containing all metadata extracted for each slide</returns>
        /// <exception cref="SlideProcessingException">
        /// This method will throw a SlideProcessingException, containing the relevant FailureReason upon any
        /// processing failure. The SlideProcessingException may contain inner exceptions.
        /// </exception>
        Task<List<Slide>> ExtractSlidesAsync(
            string inputFilePath,
            string outputFolderPath,
            CancellationToken token);
    }
}