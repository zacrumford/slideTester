using System.Collections.Generic;
using System.Linq;
using SlideTester.Common.Extensions;

namespace SlideTester.Driver;

/// <summary>
/// Class to hold the metadata extracted from a single Powerpoint slide 
/// </summary>
public class Slide
{
    #region Members and properties
    
    public string ImagePath { get; }
    public int SlideNumber { get; }
    public List<string> TitleTextFrames { get; }
    public string Title => string.Join(" ", this.TitleTextFrames);

    public string BestAvailableTitle =>
        TitleTextFrames.FirstOrDefault() ?? SubtitleTextFrames.FirstOrDefault() ?? string.Empty;
    public List<string> SubtitleTextFrames { get; }
    public string Subtitle => string.Join(" ", this.SubtitleTextFrames);
    public List<string> HeaderTextFrames { get; }
    public string HeaderText => string.Join(" ", this.HeaderTextFrames);
    public List<string> FooterTextFrames { get; }
    public string FooterText => string.Join(" ", this.FooterTextFrames);
    public List<string> BodyTextFrames { get; }
    public string BodyText => string.Join(" ", this.BodyTextFrames);
    public List<string> AnimationTextFrames { get; }
    public string AnimationText => string.Join(" ", this.AnimationTextFrames);
    
    public List<string> PresenterNotesTextFrames { get; }
    public string PresenterNotes => string.Join(" ", this.PresenterNotesTextFrames);
    public List<string> OtherTextFrames { get; }
    public string OtherText => string.Join(" ", this.OtherTextFrames);

    /// <summary>
    /// Property containing all non-title text content, separated by a " ".
    /// This is useful as it is the data format that the legacy slide handler expects
    /// </summary>
    public string Content => string.Join(separator: " ", this.AllTextFrames);

    /// <summary>
    /// Property containing all non-title text frames extracted from the slide
    /// This is useful as it is the data format that the legacy slide handler expects
    /// </summary>
    public IEnumerable<string> AllTextFrames =>
        this.TitleTextFrames.Concatenate(
            this.SubtitleTextFrames,
            this.HeaderTextFrames,
            this.FooterTextFrames,
            this.BodyTextFrames,
            this.AnimationTextFrames,
            this.OtherTextFrames,
            this.PresenterNotesTextFrames);
    
    #endregion
    
    /// <summary>
    /// Initial values ctor
    /// </summary>
    public Slide(
        int slideNumber,
        string imagePath,
        IEnumerable<string> title,
        IEnumerable<string> subtitle,
        IEnumerable<string> headers,
        IEnumerable<string> footers,
        IEnumerable<string> bodyText,
        IEnumerable<string> animationText,
        IEnumerable<string> presenterNotes,
        IEnumerable<string> otherText)
    {
        this.SlideNumber = slideNumber;
        this.ImagePath = imagePath;
        this.TitleTextFrames  = new List<string>(title);
        this.SubtitleTextFrames = new List<string>(subtitle);
        this.HeaderTextFrames = new List<string>(headers);
        this.FooterTextFrames = new List<string>(footers);
        this.BodyTextFrames = new List<string>(bodyText);
        this.AnimationTextFrames = new List<string>(animationText);
        this.PresenterNotesTextFrames = new List<string>(presenterNotes);
        this.OtherTextFrames = new List<string>(otherText);
    }
}
