using System.Collections.Generic;
using System.Drawing;

namespace SlideTester.Driver
{
    internal static class MaxPixelCountHelper
    {
        public static Size Ratio1_1 = new Size(1, 1);
        public static Size Ratio1_2 = new Size(1, 2);
        public static Size Ratio2_1 = new Size(2, 1);
        public static Size Ratio2_3 = new Size(2, 3);
        public static Size Ratio3_2 = new Size(3, 2);
        public static Size Ratio3_4 = new Size(3, 4);
        public static Size Ratio3_5 = new Size(3, 5);
        public static Size Ratio4_3 = new Size(4, 3);
        public static Size Ratio4_5 = new Size(4, 5);
        public static Size Ratio5_4 = new Size(5, 4);
        public static Size Ratio5_3 = new Size(5, 3);
        public static Size Ratio5_8 = new Size(5, 8);
        public static Size Ratio5_16 = new Size(5, 16);
        public static Size Ratio8_5 = new Size(8, 5);
        public static Size Ratio9_16 = new Size(9, 16);
        public static Size Ratio9_17 = new Size(9, 17);
        public static Size Ratio9_21 = new Size(9, 21);
        public static Size Ratio16_5 = new Size(16, 5);
        public static Size Ratio16_9 = new Size(16, 9);
        public static Size Ratio16_25 = new Size(16, 25);
        public static Size Ratio17_9 = new Size(17, 9);
        public static Size Ratio21_9 = new Size(21, 9);
        public static Size Ratio25_16 = new Size(25, 16);
        
        // AR pixel counts come from here https://www.vaughns-1-pagers.com/computer/video-resolution.htm
        // and here https://www.vaughns-1-pagers.com/computer/video-resolution.htm
        // I picked values all with pixel count close to what we originally had in PPTImageMaxPixelCount
        public static readonly Dictionary<Size, uint> MaxPixelCounts = new Dictionary<Size, uint>()
        {
            { Ratio16_25, 2400 * 1536 },
            { Ratio25_16, 2400 * 1536 },
            { Ratio16_5, 1920 * 600 },
            { Ratio5_16, 1920 * 600 },
            { Ratio16_9, 1920 * 1080 },
            { Ratio9_16, 1920 * 1080 },
            { Ratio17_9, 2048 * 1080 },
            { Ratio9_17, 2048 * 1080 },
            { Ratio1_1, 1440 * 1440 },
            { Ratio1_2, 2400 * 1200 },
            { Ratio2_1, 2400 * 1200 },
            { Ratio2_3, 1920 * 1280 },
            { Ratio3_2, 1920 * 1280 },
            { Ratio21_9, 2560 * 1080 },
            { Ratio9_21, 2560 * 1080 },
            { Ratio3_4, 1600 * 1200 },
            { Ratio3_5, 1920 * 1152 },
            { Ratio5_3, 1920 * 1152 },
            { Ratio4_3, 1600 * 1200 },
            { Ratio4_5, 1920 * 1536 },
            { Ratio5_4, 1920 * 1536 },
            { Ratio5_8, 1920 * 1200 },
            { Ratio8_5, 1920 * 1200 },
        };
    }
}
