using System;
using System.Collections.Generic;
using System.Linq;

using Aspose.Slides;

using SlideTester.Common.Extensions;

namespace SlideTester.Driver.Aspose
{
    /// <summary>
    /// Extensions to the Aspose.Slides.Slide object for ease of processing 
    /// </summary>
    internal static class AsposeSlideExtensions
    {
        /// <summary>
        /// Returns a string containing all presentation notes on a slide.
        /// Will return string.Empty if no presentation notes were found. 
        /// </summary>
        public static string PresenterNotes(this ISlide slide)
        {
            string presenterNotes = String.Empty;

            INotesSlideManager notesManager = slide.NotesSlideManager;
            INotesSlide notesSlide = notesManager?.NotesSlide;
            presenterNotes = notesSlide?.NotesTextFrame?.Text ?? string.Empty;

            return presenterNotes;
        }

        /// <summary>
        /// Returns a list of strings containing all titles on a slide.
        /// Will return an empty list if no titles were found. 
        /// </summary>
        /// 
        public static List<string> TitleText(this ISlide slide)
            => slide.GetTextForType(PlaceholderType.Title)
                .Concat(slide.GetTextForType(PlaceholderType.CenteredTitle))
                .ToList();
        
        /// <summary>
        /// Returns a list of strings containing all subtitles on a slide.
        /// Will return an empty list if no subtitles were found. 
        /// </summary>
        public static List<string> SubtitleText(
            this ISlide slide) 
            => slide.GetTextForType(PlaceholderType.Subtitle);

        /// <summary>
        /// Returns a list of strings containing all animated body text on a slide.
        /// Will return an empty list if no such text was found. 
        /// </summary>
        public static List<string> BodyAnimatedText(this ISlide slide)
            => slide.Shapes
                .Where(s => s.IsBody() && s.AnimationText().Any())
                .OrderBy(s => s.ZOrderPosition)
                .Select(s => s.AnimationText())
                .Combine()
                .ToList();
            
        /// <summary>
        /// Returns a list of strings containing all non-animated body text on a slide.
        /// Will return an empty list if no such text was found. 
        /// </summary>
        public static List<string> BodyStaticText(this ISlide slide)
            => slide.Shapes
                .Where(s => s.IsBody() && s.StaticText().Any())
                .OrderBy(s => s.ZOrderPosition)
                .Select(s => s.StaticText())
                .Combine()
                .ToList();
        
        /// <summary>
        /// Returns a list of strings containing all header text on a slide.
        /// Will return an empty list if no such text was found. 
        /// </summary>
        public static List<string> HeaderText(this ISlide slide)
            => slide.GetTextForType(PlaceholderType.Header);

        /// <summary>
        /// Returns a list of strings containing all footer text on a slide.
        /// Will return an empty list if no such text was found. 
        /// </summary>
        public static List<string> FooterText(this ISlide slide)
            => slide.GetTextForType(PlaceholderType.Footer);
        
        /// <summary>
        /// Returns a list of strings containing all non-header, non-footer, non-body, non-title and non-subtitle
        /// text on a slide. This may include things like text on a chart or table. 
        /// Will return an empty list if no such text was found. 
        /// </summary>
        public static List<string> GetTextFromNonTextBoxes(this ISlide slide)
            => slide.Shapes
                .Where(s => s.IsNonTextBoxType() && !string.IsNullOrWhiteSpace(s.AllText()))
                .Select(s => s.AllText())
                .ToList();

        /// <summary>
        /// Returns a list of strings containing all text contained in a specified shape type
        /// Will return an empty list if no such text was found. 
        /// </summary>
        private static List<string> GetTextForType(
            this ISlide slide,
            PlaceholderType placeholderType)
            => slide.Shapes
                .Where(s => s.IsType(placeholderType) && !string.IsNullOrWhiteSpace(s.AllText()))
                .Select(s => s.AllText())
                .ToList();
    }
}
