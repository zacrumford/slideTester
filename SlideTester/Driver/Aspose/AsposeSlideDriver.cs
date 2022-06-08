using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

using Aspose.Slides;
using Aspose.Slides.Export;
using ImageMagick;
using Polly;

using SlideTester.Common;
using SlideTester.Common.Extensions;
using SlideTester.Common.Log;
using PanoptoSlide = SlideTester.Driver.Slide;
using AsposeSlide = Aspose.Slides.Slide;
using AsposeLicense = Aspose.Slides.License;
using Encoder = System.Text.Encoder;

namespace SlideTester.Driver.Aspose
{
    /// <summary>
    /// ISlideDriver concrete instance designed to use Aspose.Slide NET in order to
    /// extract slide text and images from .ppt and .pptx files
    /// </summary>
    internal class AsposeSlideDriver : SafeDisposable, ISlideDriver
    {
        #region Members and properties
        
        /// <summary>
        /// Property to get the directory in which the current module exists.
        /// Useful when adding custom font directories
        /// </summary>
        private static Lazy<string> CurrentDirectory { get; } = new Lazy<string>(
            () => Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location));
        
        /// <summary>
        /// Scratch space with which this object instance will work
        /// </summary>
        private ScratchSpace Scratch { get; }
        
        /// <summary>
        /// Member to ensure that we only perform global setup operations a single time
        /// </summary>
        private static bool GlobalSetupComplete { get; set; } = false;
        
        /// <summary>
        /// Lock to ensure we don't attempt global setup operations concurrently
        /// </summary>
        private static object GlobalSetupLock { get; } = new object();
        
        /// <summary>
        /// Bulkhead policy to control the max parallelism of Async calls to the aspose engine
        /// </summary>
        private AsyncPolicy SlideProcessingBulkhead { get; }

        #endregion
        
        
        #region public methods
        
        /// <summary>
        /// Initial values ctor.
        /// The first object created will execute Aspose GlobalSetup
        /// </summary>
        public AsposeSlideDriver(
            int maxParallelization)
        {
            AsposeSlideDriver.GlobalSetup(
                Settings.Default.AsposeLicense,
                Settings.Default.CustomFontDirectories);
            
            this.Scratch = new ScratchSpace();
            this.SlideProcessingBulkhead =  Policy.BulkheadAsync(
                maxParallelization,
                maxQueuingActions: int.MaxValue);
        }
        
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
        public async Task<List<PanoptoSlide>> ExtractSlidesAsync(
            string inputFilePath,
            string outputFolderPath,
            CancellationToken token)
        {
            List<PanoptoSlide> results = null;
            
            ChkArg.IsNotNull(token, nameof(token));
            ChkArg.IsTrue(File.Exists(inputFilePath), nameof(inputFilePath));
            ChkArg.IsTrue(Directory.Exists(outputFolderPath), nameof(outputFolderPath));
            token.ThrowIfCancellationRequested();

            try
            {
                using Presentation presentation = new Presentation(inputFilePath);

                // Note: The presentation should fail to open if password protected/encrypted.
                // We shouldn't get this far. Instead, InvalidPasswordException should be thrown from above.
                ChkExpect.IsFalse(
                    presentation.ProtectionManager.IsEncrypted,
                    "Encrypted presentations are not supported");

                // Note, we round up on sub-pixels
                Size originalSize = new Size(
                    (int) (presentation.SlideSize.Size.Width + 0.5f),
                    (int) (presentation.SlideSize.Size.Height + 0.5f));
                Size scaledSize = originalSize.ComputeScaledDimensions(
                    ResolutionHelper.MaxPixelCount(originalSize));

                var tasks = new List<Task<PanoptoSlide>>();

                foreach (ISlide slide in presentation.Slides)
                {
                    // Use a bulkhead for processing so we don't attempt to do upwards of 1000 slide
                    // extraction operations concurrently.
                    tasks.Add(this.SlideProcessingBulkhead.ExecuteAsync(async () =>
                        await this.ProcessSlide(outputFolderPath, scaledSize, slide, token).ConfigureAwait(false)));
                }

                results = (await Task.WhenAll(tasks).ConfigureAwait(false))
                    .OrderBy(s => s.SlideNumber)
                    .ToList();
            }
            catch (InvalidPasswordException ex)
            {
                // An Aspose.InvalidPasswordException will be thrown if the file is password protected.
                // Grab the exception, wrap it in our SlideProcessingException with the correct failure reason
                Logger.Write("Slide deck is PW protected", ex);
                var slideProcessingException = new SlideProcessingException(
                    inputFilePath,
                    SlideProcessingException.FailureReason.PasswordProtected,
                    ex);
                throw slideProcessingException;
            }
            catch (SlideProcessingException ex)
            {
                Logger.Write("Unexpected SlideProcessingException", ex);
                throw;
            }
            catch (Exception ex)
            {
                // We are not an Exception of type SlideProcessingException. We need to wrap in an 
                // SlideProcessingException exception, log and throw;
                Logger.Write("Unexpected Exception during slide processing", ex);
                var slideProcessingException = new SlideProcessingException(
                    inputFilePath,
                    SlideProcessingException.FailureReason.FailureUnknown,
                    ex);
                throw slideProcessingException;
            }

            return results;
        }

        #endregion
        
        
        #region private/protected methods

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

        /// <summary>
        /// Processes a single slide. Extracts all textual information on slide and exports slide as image
        /// to output path, with specified image size, as a .jpg 
        /// </summary>
        /// <returns>A Slide object containing all extracted text data and the slide image path</returns>
        private async Task<PanoptoSlide> ProcessSlide(
            string outputFolderPath,
            Size desiredImageSize,
            ISlide slide,
            CancellationToken token)
        {
            ChkArg.IsNotNull(token, nameof(token));
            ChkArg.IsTrue(Directory.Exists(outputFolderPath), nameof(outputFolderPath));
            token.ThrowIfCancellationRequested();
            
            string svgPath = Path.Combine(this.Scratch.UniqueFilePath(".svg"));
            ChkExpect.IsFalse(File.Exists(svgPath), nameof(svgPath));
            List<string> title = slide.TitleText();
            List<string> animationText = slide.BodyAnimatedText();
            List<string> bodyText = slide.BodyStaticText();
            List<string> subTitle = slide.SubtitleText();
            List<string> presenterNotes = new List<string>(){ slide.PresenterNotes()}; // Aspose provides a single string
            List<string> otherText = slide.GetTextFromNonTextBoxes();
            List<string> headerText = slide.HeaderText();
            List<string> footerText = slide.FooterText();

            // Export slide as bmp from Aspose, load into ImageMagick image
            using Bitmap bitmap = slide.GetThumbnail(1.0f, 1.0f);
            IMagickImage slideImage;
            MagickFactory f = new MagickFactory();
            using (MemoryStream ms = new MemoryStream())
            {
                bitmap.Save(ms, ImageFormat.Bmp);
                ms.Position = 0;
                slideImage = new MagickImage(await f.Image.CreateAsync(ms, token));
            }
            
            string pngPath = Path.Combine(outputFolderPath, $"{slide.SlideNumber}-{Guid.NewGuid()}.png");
            
            // Use ImageMagick to convert our bitmap to thw desired output format (png), resize as needed
            if (slideImage.Width != desiredImageSize.Width || slideImage.Height != desiredImageSize.Height)
            {
                slideImage.Resize(desiredImageSize.Width, desiredImageSize.Height);
            }

            await using Stream pngStream = File.Create(pngPath);
            await slideImage.WriteAsync(pngStream, MagickFormat.Png, token).ConfigureAwait(false);
            
            return new PanoptoSlide(
                slide.SlideNumber,
                pngPath,
                title,
                subTitle,
                headerText,
                footerText,
                bodyText,
                animationText,
                presenterNotes,
                otherText);
        }

        private static ImageCodecInfo GetJpegEncoderInfo()
            => GetEncoderInfo("image/jpeg");
        
        private static ImageCodecInfo GetEncoderInfo(String mimeType)
            => ImageCodecInfo.GetImageEncoders()
                .FirstOrDefault(e => string.Equals(e.MimeType, mimeType, StringComparison.InvariantCultureIgnoreCase) );
        
        /// <summary>
        /// Method to perform global setup (licensing and custom font directory configuration)  
        /// </summary>
        private static void GlobalSetup(
            string licenseBlob,
            IEnumerable<string> customFontDirectories,
            bool forceReinitialization = false)
        {
            lock (GlobalSetupLock)
            {
                if (!GlobalSetupComplete || forceReinitialization)
                {
                    var license = new AsposeLicense();
                    
                    using Stream licenseStream = licenseBlob.ToStream(Encoding.UTF8);
                    license.SetLicense(licenseStream);
                    
                    // We want to hard fail if we are not licensed. Else we will "succeed" processing but
                    // with bad data (truncated text and images that have Aspose licensing warnings) 
                    ChkExpect.IsTrue(license.IsLicensed(), "Aspose license was not loaded");
                    
                    // Load custom fonts from our input dirs (if specified) and our current dir
                    var customFontsDirs = new List<string>
                    {
                        CurrentDirectory.Value
                    };
                    customFontsDirs.AddRange(customFontDirectories);
                    FontsLoader.LoadExternalFonts(customFontsDirs.ToArray());
                    
                    GlobalSetupComplete = true;
                }
            }
        }
        
        #endregion
    }
}