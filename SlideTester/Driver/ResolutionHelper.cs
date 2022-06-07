using System;
using System.Collections.Generic;
using System.Drawing;

using SlideTester.Common;
using SlideTester.Common.Extensions;

namespace SlideTester.Driver
{
    internal static class ResolutionHelper
    {        // AR pixel counts come from here https://www.vaughns-1-pagers.com/computer/video-resolution.htm
        // and here https://www.vaughns-1-pagers.com/computer/video-resolution.htm
        // I picked values all with pixel count close to what we originally had in PPTImageMaxPixelCount
        private static Lazy<Dictionary<Size, uint>> MaxPixelCounts { get; } =
            new Lazy<Dictionary<Size, uint>>(() => 
                new Dictionary<Size, uint>()
                {
                    { AspectRatio.Ratio16_25, 2400 * 1536 },
                    { AspectRatio.Ratio25_16, 2400 * 1536 },
                    { AspectRatio.Ratio16_5, 1920 * 600 },
                    { AspectRatio.Ratio5_16, 1920 * 600 },
                    { AspectRatio.Ratio16_9, 1920 * 1080 },
                    { AspectRatio.Ratio9_16, 1920 * 1080 },
                    { AspectRatio.Ratio17_9, 2048 * 1080 },
                    { AspectRatio.Ratio9_17, 2048 * 1080 },
                    { AspectRatio.Ratio1_1, 1440 * 1440 },
                    { AspectRatio.Ratio1_2, 2400 * 1200 },
                    { AspectRatio.Ratio2_1, 2400 * 1200 },
                    { AspectRatio.Ratio2_3, 1920 * 1280 },
                    { AspectRatio.Ratio3_2, 1920 * 1280 },
                    { AspectRatio.Ratio21_9, 2560 * 1080 },
                    { AspectRatio.Ratio9_21, 2560 * 1080 },
                    { AspectRatio.Ratio3_4, 1600 * 1200 },
                    { AspectRatio.Ratio3_5, 1920 * 1152 },
                    { AspectRatio.Ratio5_3, 1920 * 1152 },
                    { AspectRatio.Ratio4_3, 1600 * 1200 },
                    { AspectRatio.Ratio4_5, 1920 * 1536 },
                    { AspectRatio.Ratio5_4, 1920 * 1536 },
                    { AspectRatio.Ratio5_8, 1920 * 1200 },
                    { AspectRatio.Ratio8_5, 1920 * 1200 },
                });
        
        public static uint MaxPixelCount(Size size)
        {
            uint maxPixelCount;

            Size aspectRatio = size.GetAspectRatio();

            if (MaxPixelCounts.Value.ContainsKey(aspectRatio))
            {
                maxPixelCount = MaxPixelCounts.Value[aspectRatio];
            }
            else
            {
                maxPixelCount = Settings.Default.SlideMaxPixelCountFallback;
            }

            return maxPixelCount;
        }
    }
}