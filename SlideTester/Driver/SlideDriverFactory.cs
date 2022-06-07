using System;
using Panopto.Common;
using Panopto.Common.Media.Slides;

namespace SlideTester.Driver
{
    public class SlideDriverFactory
    {
        public static SlideDriverFactory Instance => LazyInstance.Value;
        private static Lazy<SlideDriverFactory> LazyInstance { get; } = new Lazy<SlideDriverFactory>(
            () => new SlideDriverFactory());

        public virtual ISlideDriver CreateDriver(
            SlideDriverType slideDriverType,
            int? maxParallelization = null)
        {
            ISlideDriver result = null;
            
            switch (slideDriverType)
            {
                case SlideDriverType.Aspose:
                    result = new Aspose.AsposeSlideDriver(
                        maxParallelization ?? Settings.Default.MaxSlideProcessingParallelization);
                    break;
                case SlideDriverType.Powerpoint:
                    result = new Powerpoint.PowerpointSlideDriver();
                    break;
                default:
                    throw new Exception($"Unsupported value for {nameof(slideDriverType)}: {slideDriverType}" );
                    break;
            }

            return result;
        }
    }
}