using System.Collections.Generic;
using System.Linq;
using Aspose.Slides;
using Aspose.Slides.Animation;

namespace SlideTester.Driver.Aspose
{
    /// <summary>
    /// Extensions to the Aspose.IShape interface 
    /// </summary>
    internal static class AsposeShapeExtensions
    {
        /// <summary>
        /// Returns true if the shape is a title text-box 
        /// </summary>
        public static bool IsTitle(this IShape shape)
            => shape.IsType(PlaceholderType.CenteredTitle) || shape.IsType(PlaceholderType.Title);
        
        /// <summary>
        /// Returns true if the shape is a subtitle text-box 
        /// </summary>
        public static bool IsSubTitle(this IShape shape) 
            => shape.IsType(PlaceholderType.Subtitle);
        
        /// <summary>
        /// Returns true if the shape is a body text-box 
        /// </summary>
        public static bool IsBody(this IShape shape) 
            => shape.IsType(PlaceholderType.Body);
        
        /// <summary>
        /// Returns true if the shape is a footer text-box 
        /// </summary>
        public static bool IsFooter(this IShape shape) 
            => shape.IsType(PlaceholderType.Footer);
            
        /// <summary>
        /// Returns true if the shape is a header text-box 
        /// </summary>
        public static bool IsHeader(this IShape shape) 
            => shape.IsType(PlaceholderType.Header);
        
        /// <summary>
        /// Returns true if the shape is a slide number text-box 
        /// </summary>
        public static bool IsSlideNumber(this IShape shape) 
            => shape.IsType(PlaceholderType.SlideNumber);


        /// <summary>
        /// Returns true if the shape is a slide DateTime text-box 
        /// </summary>
        public static bool IsDateAndTime(this IShape shape) 
            => shape.IsType(PlaceholderType.DateAndTime);
        
        /// <summary>
        /// Returns true if the shape is not a known text box type 
        /// </summary>
        public static bool IsNonTextBoxType(this IShape shape)
            => !shape.IsBody()
               && !shape.IsFooter()
               && !shape.IsHeader()
               && !shape.IsTitle()
               && !shape.IsSubTitle()
               && !shape.IsSlideNumber()
               && !shape.IsDateAndTime();

        /// <summary>
        /// Gets the ITextFrame from a shape, if one exists on the shape. Else returns null.
        /// </summary>
        public static ITextFrame GetTextFrame(
            this IShape shape)
        {
            // Some times a shape may be a ITextFrame itself or it may be an AutoShape 
            // which in turn contains an ITextFrame. Handle both cases. If neither case is
            // true then we return null
            // ReSharper disable once SuspiciousTypeConversion.Global
            if (!(shape is ITextFrame textFrame))
            {
                // Try a different way to get the text frame
                AutoShape autoShape = shape as AutoShape;
                textFrame = autoShape?.TextFrame;
            }

            return textFrame;
        }
        
        /// <summary>
        /// Returns true if the shape type is equal to the specified type, else false
        /// </summary>
        public static bool IsType(
            this IShape shape,
            PlaceholderType placeholderType)
        {
            bool result = false;
            IPlaceholder placeholder = shape.Placeholder;

            if (placeholder != null && placeholder.Type == placeholderType)
            {
                result = true;
            }

            return result;
        }

        /// <summary>
        /// Returns true if the shape has animations, else false
        /// </summary>
        public static bool HasAnimation(this IShape shape)
        {
            bool result = false;
            ISequence sequence = shape.Slide.Timeline.MainSequence;
            ITextFrame textFrame = shape.GetTextFrame();
            
            if (textFrame != null)
            {
                foreach (IParagraph paragraph in textFrame.Paragraphs)
                {
                    IEffect[] effects = sequence.GetEffectsByParagraph(paragraph);

                    if (effects.Any())
                    {
                        result = true;
                        break;
                    }
                }
            }

            return result;
        }
        
        /// <summary>
        /// Returns a list of all text in a shape which is contained within animations.
        /// All text returned will be in animation sorted order
        /// If no animations exist or if no text is associated it animations, an empty list will be returned.
        /// </summary>
        public static List<string> AnimationText(this IShape shape)
        {
            var animationText = new List<string>();
            ISequence sequence = shape.Slide.Timeline?.MainSequence;
            ITextFrame textFrame = shape.GetTextFrame();
            
            if (textFrame != null && sequence != null)
            {
                // Note: Animations on paragraphs always are in paragraph order, so enumerating 
                // in this way always produces ordered animations
                foreach (IParagraph paragraph in textFrame.Paragraphs ?? Enumerable.Empty<IParagraph>())
                {
                    IEffect[] effects = sequence.GetEffectsByParagraph(paragraph);

                    if (effects.Any() && !string.IsNullOrWhiteSpace(paragraph.Text))
                    {
                        animationText.Add(paragraph.Text);
                    }
                }
            }

            return animationText;
        }
        
        /// <summary>
        /// Returns a list of all text in a shape which is not contained within animations.
        /// </summary>
        public static List<string> StaticText(this IShape shape)
        {
            var staticText = new List<string>();
            ISequence sequence = shape.Slide.Timeline?.MainSequence;
            ITextFrame textFrame = shape.GetTextFrame();

            if (textFrame != null && sequence != null)
            {
                foreach (IParagraph paragraph in textFrame.Paragraphs ?? Enumerable.Empty<IParagraph>())
                {
                    IEffect[] effects = sequence.GetEffectsByParagraph(paragraph);

                    if (!effects.Any() && !string.IsNullOrWhiteSpace(paragraph.Text))
                    {
                        staticText.Add(paragraph.Text);
                    }
                }
            }
            else if (textFrame != null)
            {
                staticText.Add(shape.AllText());
            }

            return staticText;
        }

        /// <summary>
        /// Returns a list of all text in a shape
        /// </summary>
        public static string AllText(this IShape shape)
            => shape.GetTextFrame()?.Text ?? string.Empty;
    }
}