using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;

using Amazon;
using Amazon.S3;
using Amazon.S3.Model;
using Amazon.S3.Transfer;
using ImageMagick;

using Panopto.Common.Media.Slides;
using SlideTester.App.CommandLine;
using SlideTester.Common;
using SlideTester.Common.Log;
using SlideTester.Driver;

namespace SlideTester.App
{
    internal static class ImageCompareThresholds
    {
        public static double Error { get; } = 0.1;
        public static double Warning { get; } = 0.05;
        public static double Info { get; } = 0.01;
    }
    
    internal abstract class BaseProcessor<T> : SafeDisposable, IVerbProcessor where T : BaseOptions
    {
        private static List<SlideDriverType> DriverTypesToTest { get; } = new List<SlideDriverType>()
        {
            SlideDriverType.Aspose,
            SlideDriverType.Powerpoint,
        };
        
        protected T Options { get; }
        private ScratchSpace Scratch { get; }

        protected BaseProcessor(T options)
        {
            this.Options = options;
            this.Scratch = new ScratchSpace();
            if (!Directory.Exists(this.Options.OutputPath))
            {
                Directory.CreateDirectory(options.OutputPath);
            }

            this.ReconfigureApplicationLogging();
        }
        
        public abstract Task DoWork(CancellationToken token);

        /// <summary>
        /// Method to clean up managed resources.  This method will be invoked 
        /// once and only once, and only if there is an explicit dispose.
        /// </summary>
        protected override void CleanupDisposableObjects()
        {
            SafeDisposable.Dispose(Scratch);
        }

        /// <summary>
        /// Method to clean up native resources.  This method will be invoked
        /// once and only once, regardless of if there is an explicit dispose.
        /// </summary>
        protected override void CleanupUnmanagedResources()
        {
            // NOP
        }

        private async Task CompareOutput(
            DateTime startTime,
            string powerpointFileUri,
            List<(SlideDriverType slideDriverType, List<Slide> slides)> results,
            CancellationToken token)
        {
            string parentFolderPath = Path.Combine(this.Options.OutputPath, Path.GetFileName(powerpointFileUri));
            string compareLogPath = Path.Combine(parentFolderPath, $"{startTime.Ticks}_results.txt");
            if (!Directory.Exists(parentFolderPath))
            {
                Directory.CreateDirectory(parentFolderPath);
            }
            StringBuilder logMessage = new StringBuilder();
            await using StreamWriter compareResultsFileStream = new StreamWriter(compareLogPath);
            await using StreamWriter consoleOutputStream = new StreamWriter(Console.OpenStandardOutput())
            {
                AutoFlush = true
            };
            List<StreamWriter> streamWriters = new List<StreamWriter>()
            {
                compareResultsFileStream,
            };

            if (this.Options.ShouldPrintCompareResults)
            {
                streamWriters.Add(consoleOutputStream);
            }
            
            ChkArg.AreEqual(2, results.Count, "Only two slide processors are currently supported");
            List<Slide> asposeSlides = results.First(r => r.slideDriverType == SlideDriverType.Aspose).slides;
            List<Slide> pptSlides = results.First(r => r.slideDriverType == SlideDriverType.Powerpoint).slides;
            
            if (pptSlides.Count != asposeSlides.Count)
            {
                logMessage.AppendLine(
                    $"[Error]: Unexpected slide count.\nAspose: {asposeSlides.Count} \nPpt: {pptSlides.Count}\nFile:{powerpointFileUri}\n");
            }

            for (int i = 0; i < pptSlides.Count; ++i)
            {
                if (pptSlides[i].SlideNumber != asposeSlides[i].SlideNumber)
                {
                    logMessage.AppendLine(
                        $"[Error]: Unexpected slide number.\nAspose: {asposeSlides[i].SlideNumber}\nPpt: {pptSlides[i].SlideNumber}\nFile:{powerpointFileUri}\n");
                }
                
                if (!AreEqualSansWhitespace(pptSlides[i].BestAvailableTitle, asposeSlides[i].BestAvailableTitle))
                {
                    logMessage.AppendLine(
                        $"[Error]: Unexpected title. \nAspose: {asposeSlides[i].BestAvailableTitle}\nPpt: {pptSlides[i].BestAvailableTitle}\nFile:{powerpointFileUri}\n");
                }
                
                if (!AreEqualIgnoreOrder(pptSlides[i].AllTextFrames, asposeSlides[i].AllTextFrames))
                {
                    logMessage.AppendLine(
                        $"[Error]: Unexpected content. \nAspose: {asposeSlides[i].Content}\nPpt: {pptSlides[i].Content}\nFile:{powerpointFileUri}\n");
                }

                if (!this.Options.SkipImageCompare)
                {
                    string diffImagePath = Path.Combine(
                        parentFolderPath,
                        $"{startTime.Ticks}_slide{pptSlides[i].SlideNumber}.jpg");
                    string imageCompareLog = await this.CompareImages(
                        diffImagePath,
                        asposeSlides[i].ImagePath,
                        pptSlides[i].ImagePath,
                        token).ConfigureAwait(false);
                    if (!string.IsNullOrWhiteSpace(imageCompareLog))
                    {
                        logMessage.AppendLine(imageCompareLog);
                    }
                }
            }
            
            foreach (var streamWriter in streamWriters)
            {
                await streamWriter.WriteLineAsync(logMessage, token).ConfigureAwait(false);
            }
        }

        private bool AreEqualIgnoreOrder(
            IEnumerable<string> lhs,
            IEnumerable<string> rhs)
        {
            string orderedLhs = string.Join(
                " ",
                lhs.Where(str => !string.IsNullOrWhiteSpace(str)).OrderBy(str => str));
            string orderedRhs = string.Join(
                " ",
                rhs.Where(str => !string.IsNullOrWhiteSpace(str)).OrderBy(str => str));
            
            return this.AreEqualSansWhitespace(orderedLhs, orderedRhs);
        }
        
        // Issue #1: White space differences
        private bool AreEqualSansWhitespace(
            string str1,
            string str2)
        {
            string normalized1 = Regex.Replace(str1, @"\s", "");
            string normalized2 = Regex.Replace(str2, @"\s", "");

            return string.Equals(
                normalized1, 
                normalized2, 
                StringComparison.InvariantCulture);
        }
        
        private async Task<string> CompareImages(
            string outputDiffImagePath,
            string referenceImagePath,
            string comparandImagePath,
            CancellationToken token)
        {
            StringBuilder logMessage = new StringBuilder();
            
            using var referenceImage = new MagickImage(referenceImagePath);
            using var comparandImage = new MagickImage(comparandImagePath);

            if (referenceImage.Width != comparandImage.Width)
            {
                logMessage.AppendLine(
                    $"[ERROR] Mismatching width: {referenceImagePath}:{referenceImage.Width} vs. {comparandImagePath}:{comparandImage.Width}");
            }
            
            if (referenceImage.Height != comparandImage.Height)
            {
                logMessage.AppendLine(
                    $"[ERROR] Mismatching height: {referenceImagePath}:{referenceImage.Height} vs. {comparandImagePath}:{comparandImage.Height}");
            }
            
            if (referenceImage.Format != comparandImage.Format)
            {
                logMessage.AppendLine(
                    $"[ERROR] Mismatching format: {referenceImagePath}:{referenceImage.Format} vs. {comparandImagePath}:{comparandImage.Format}");
            }
            
            // Next do an image compare
            var imgDiff = new MagickImage();
            double diff = referenceImage.Compare(comparandImage, ErrorMetric.RootMeanSquared, imgDiff);
            //await imgDiff.WriteAsync(outputDiffImagePath).ConfigureAwait(false);
            if (diff >= ImageCompareThresholds.Error)
            {
                logMessage.AppendLine(
                    $"[ERROR] Images differ too much ({diff}): {referenceImagePath} vs. {comparandImagePath}");
            }
            else if (diff < ImageCompareThresholds.Error && diff >= ImageCompareThresholds.Warning)
            {
                logMessage.AppendLine(
                    $"[WARNING] Images differ too much ({diff}): {referenceImagePath} vs. {comparandImagePath}");
            }
            else if (diff < ImageCompareThresholds.Warning && diff >= ImageCompareThresholds.Info)
            {
                logMessage.AppendLine(
                    $"[INFO] Images differ tool much ({diff}): {referenceImagePath} vs. {comparandImagePath}");
            }
            
            using var imageCollection = new MagickImageCollection();
            imageCollection.Add(referenceImage);
            imageCollection.Add(comparandImage);
            using var comparandImageClone = comparandImage.Clone();
            imageCollection.Add(comparandImageClone);
            imageCollection.Add(imgDiff);
            
            var montageSettings = new MontageSettings()
            {
                BackgroundColor = MagickColors.None, // -background none
                Shadow = false, // -shadow
                Geometry = new MagickGeometry(0, 0, imgDiff.Width, imgDiff.Height) // -geometry +5+5
            };
            
            using var montageImage = imageCollection.Montage(montageSettings);
            await montageImage.WriteAsync(outputDiffImagePath, token).ConfigureAwait(false);

            return logMessage.ToString();
        }
        
        
        protected async Task ProcessFile(
            string powerpointFileUri,
            string region,
            CancellationToken token)
        {
            token.ThrowIfCancellationRequested();

            Console.WriteLine($"Starting processing of {powerpointFileUri}.");
            
            DateTime startTime = DateTime.UtcNow;

            if (powerpointFileUri.StartsWith("s3://"))
            {
                ChkExpect.IsNotNullOrWhiteSpace(region, "--region parameter must be set if uri is an s3 uri");
                ChkExpect.IsNotNull(RegionEndpoint.GetBySystemName(region), "Invalid aws region");
            }
            
            Console.WriteLine($"Copying file from {powerpointFileUri} to scratch location");
            if ( !await this.ExistsAsync(powerpointFileUri, region, token).ConfigureAwait(false))
            {
                Console.WriteLine($"[[ERROR]] File does not exist: {powerpointFileUri}.");
                return;
            }
            
            string localFilePath = await this.MakeLocalAsync(powerpointFileUri, region, token).ConfigureAwait(false);
            ChkExpect.IsTrue(File.Exists(localFilePath), "Copied file was not found. Does the source exist?");
            Console.WriteLine($"Copy operation complete. Beginning slide processing.");
            
            var taskData = new List<(SlideDriverType slideDriverType, Task<(TimeSpan duration, List<Slide> slides)> task)>();
            string filename = Path.GetFileName(powerpointFileUri);
            
            foreach (SlideDriverType driverType in DriverTypesToTest)
            {
                // To avoid access violations, make a copy of the input ppt
                string pptCopy = this.Scratch.UniqueFilePath(Path.GetExtension(localFilePath));
                File.Copy(localFilePath, pptCopy);
               
                ISlideDriver driver = SlideDriverFactory.Instance.CreateDriver(
                    driverType, 
                    this.Options.MaxSlideParallelization);
                taskData.Add(
                    (
                        driverType, 
                        Task.Run(async () =>
                            {
                                Stopwatch timer = Stopwatch.StartNew();
                                List<Slide> results = await driver.ExtractSlidesAsync(
                                    pptCopy,
                                    this.Scratch.CreateUniqueFolder(),
                                    token).ConfigureAwait(false);
                                return (timer.Elapsed, results);
                            }, 
                            token)
                    ));
            }

            try
            {
                Task.WhenAll(taskData.Select(td => td.task)).Wait(token);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[[ERROR]] Processing failure: {ex}");
            }

            TimeSpan processingTime = DateTime.UtcNow - startTime;
            
            Console.WriteLine($"Processing complete. Took {processingTime.TotalSeconds} seconds to process with all drivers.");
            Console.WriteLine($"Writing results to output");

            bool anyFailed = false;
            foreach ((SlideDriverType slideDriverType, Task<(TimeSpan duration, List<Slide> slides)> task) in taskData)
            {
                if (task.IsFaulted)
                {
                    Console.WriteLine($"[[ERROR]] {slideDriverType} failed processing for file {filename}");
                    this.WriteFailResults(
                        startTime, 
                        filename,
                        slideDriverType,
                        task.Exception);
                    anyFailed = true;
                }
                else
                {
                    this.WriteSuccessResults(
                        startTime, 
                        task.Result.duration,
                        filename,
                        slideDriverType,
                        task.Result.slides);
                }
            }

            if (this.Options.ShouldCompareOutput && !anyFailed)
            {
                await this.CompareOutput(
                    startTime,
                    powerpointFileUri,
                    taskData.Select(td => (td.slideDriverType, td.task.Result.slides)).ToList(),
                    token).ConfigureAwait(false);
            }
            
            Console.WriteLine($"Processing of {powerpointFileUri} complete");
        }

        private async Task<string> MakeLocalAsync(
            string pptUri,
            string region,
            CancellationToken token)
        {
            string localPath = this.Scratch.UniqueFilePath(Path.GetExtension(pptUri));
            
            token.ThrowIfCancellationRequested();
            
            if (pptUri.StartsWith("s3://"))
            {
                ChkExpect.IsTrue(
                    await this.TryGetS3ObjectAsync(pptUri, region, localPath, token).ConfigureAwait(false),
                    $"Failed to download s3 object: {pptUri}");
            }
            else
            {
                File.Copy(pptUri, localPath);
            }

            return localPath;
        }

        private async Task<bool> ExistsAsync(
            string uri,
            string region,
            CancellationToken token)
        {
            bool exists;
            
            if (uri.StartsWith("s3://"))
            {
                exists = await this.S3ObjectExistsAsync(uri, region, token).ConfigureAwait(false);
            }
            else
            {
                exists = File.Exists(uri);
            }

            return exists;
        }
        
        private void WriteSuccessResults(
            DateTime startTime,
            TimeSpan processingDuration,
            string pptFileName,
            SlideDriverType slideDriverType,
            IEnumerable<Slide> results)
        {
            string outFolder = Path.Combine(
                this.Options.OutputPath,
                $"{pptFileName}\\{startTime.Ticks}_{slideDriverType}");
            if (!Directory.Exists(outFolder))
            {
                Directory.CreateDirectory(outFolder);
            }
            string output = Newtonsoft.Json.JsonConvert.SerializeObject(results);
            
            File.WriteAllText(Path.Combine(outFolder, "results.json"), output);

            StringBuilder perfText = new StringBuilder();
            perfText.AppendLine($"Start time: {startTime}");
            perfText.AppendLine($"End time: {startTime + processingDuration}");
            perfText.AppendLine($"Processing duration (seconds): {processingDuration.TotalSeconds}");
            File.WriteAllText(Path.Combine(outFolder, "perfData.txt"), perfText.ToString());
            
            foreach (Slide slide in results)
            {
                File.Copy(
                    slide.ImagePath, 
                    Path.Combine(outFolder, Path.GetFileName(slide.ImagePath) ?? "unsetFilename"));
            }
        }

        private void WriteFailResults(
            DateTime startTime,
            string pptFileName,
            SlideDriverType slideDriverType,
            Exception ex)
        {
            string outFolder = Path.Combine(
                this.Options.OutputPath,
                $"{pptFileName}\\{startTime.Ticks}_{slideDriverType}");
            if (!Directory.Exists(outFolder))
            {
                Directory.CreateDirectory(outFolder);
            }

            StringBuilder perfText = new StringBuilder();
            perfText.AppendLine($"Processing failed. Exception:\n{ex}");
            perfText.AppendLine($"Start time: {startTime}");
            perfText.AppendLine($"End time: N/A");
            perfText.AppendLine($"Processing duration (seconds): N/A");
            File.WriteAllText(Path.Combine(outFolder, "perfData.txt"), perfText.ToString());
        }
        
        private async Task<bool> S3ObjectExistsAsync(
            string s3SourceUri,
            string region,
            CancellationToken token)
        {
            (string bucket, string objectKey) = this.SplitS3Uri(s3SourceUri);
            RegionEndpoint regionEndpoint = RegionEndpoint.GetBySystemName(region);
            
            ChkArg.IsNotNull(token, nameof(token));
            token.ThrowIfCancellationRequested();
            
            AmazonS3Client s3Client = new AmazonS3Client(regionEndpoint);
            GetObjectMetadataResponse response = null;

            GetObjectMetadataRequest headRequest = new GetObjectMetadataRequest()
            {
                BucketName = bucket,
                Key = objectKey
            };

            try
            {
                response = await s3Client.GetObjectMetadataAsync(headRequest, token)
                    .ConfigureAwait(false);
            }
            catch (AmazonS3Exception s3e) when (s3e.StatusCode == HttpStatusCode.NotFound)
            {
                // look for a 404 (nonexistent).  if it's anything else, don't catch.
            }
            return response != null;
        }
        
        private async Task<bool> TryGetS3ObjectAsync(
            string s3SourceUri,
            string region,
            string localDestinationPath, 
            CancellationToken token)
        {
            (string bucket, string objectKey) = this.SplitS3Uri(s3SourceUri);
            RegionEndpoint regionEndpoint = RegionEndpoint.GetBySystemName(region);
            
            ChkArg.IsNotNull(token, nameof(token));
            ChkArg.FileDoesNotExist(localDestinationPath, nameof(localDestinationPath));
            token.ThrowIfCancellationRequested();
            
            try
            {
                AmazonS3Client s3Client = new AmazonS3Client(regionEndpoint);
                using var transferClient = new TransferUtility(s3Client);
                    
                var downloadRequest = new TransferUtilityDownloadRequest()
                {
                    BucketName = bucket,
                    Key = objectKey,
                    FilePath = localDestinationPath,
                };

                await transferClient.DownloadAsync(downloadRequest, token).ConfigureAwait(false);
            }
            catch (Exception ex)
            {
                Logger.Write($"Amazon_S3_GetObjectWithoutRetry_LocalPath_Exception bucket: {bucket}, objectKey:{objectKey}", ex);

                throw;
            }

            return File.Exists(localDestinationPath);
        }
        
        private (string bucket, string objectKey) SplitS3Uri(
            string s3Uri)
        {
            var objectKeyBuilder = new StringBuilder();
            // example: s3://bucket/foo/bar/stuff.exe
            string[] split = s3Uri.Split("/");

            ChkExpect.IsTrue(split.Length >= 4, $"Invalid format for s3 uri ({s3Uri}). Expected: 's3://[bucket]/[objectKey]'");
            
            // split[0]: s3:
            // split[1]: (string.empty)
            string bucket = split[2];
            for (int i = 3; i < split.Length; ++i)
            {
                objectKeyBuilder.Append(split[i]);
                if (i + 1 < split.Length)
                {
                    // Add folder divider for all but the last objectKeyPart
                    objectKeyBuilder.Append("/");
                }
            }
            
            return (bucket, objectKeyBuilder.ToString());
        }
        
        /// <summary>
        /// Reconfigure the application logging IF NEEDED based on options.
        /// If no logging parameter were set on the commandline we perform no reconfiguration
        /// </summary>
        private void ReconfigureApplicationLogging()
        {
            // If any explicit logging config was specified then reconfigure logging for the process, else use the default
            List<ILogSinkConfiguration> logSinks = new List<ILogSinkConfiguration>()
            {
                new ConsoleLogSinkConfiguration(this.Options.ConsoleLogConfig),
                new ApplicationLogSinkConfiguration(Options.EventLogConfig),
            };

            Logger.ReconfigureLogger(logSinks.ToArray());

            // Log a message 
            Logger.Write("BackendTaskWorker_LoggerReconfigured");
        }
    }
}