using Aspose.Slides;

namespace AsposeMcpServer.Helpers.PowerPoint;

/// <summary>
///     Helper class for PowerPoint hyperlink operations.
/// </summary>
public static class PptHyperlinkHelper
{
    /// <summary>
    ///     Gets hyperlinks from a slide as JSON objects.
    ///     Detects both shape-level and portion-level (text) hyperlinks.
    /// </summary>
    /// <param name="presentation">The presentation object.</param>
    /// <param name="slide">The slide to extract hyperlinks from.</param>
    /// <returns>A list of hyperlink objects containing shape index, level, trigger type, and URL.</returns>
    public static List<object> GetHyperlinksFromSlide(IPresentation presentation, ISlide slide)
    {
        List<object> hyperlinksList = [];

        for (var shapeIndex = 0; shapeIndex < slide.Shapes.Count; shapeIndex++)
        {
            if (slide.Shapes[shapeIndex] is not IAutoShape autoShape) continue;

            AddShapeLevelHyperlinks(presentation, autoShape, shapeIndex, hyperlinksList);
            AddTextLevelHyperlinks(presentation, autoShape, shapeIndex, hyperlinksList);
        }

        return hyperlinksList;
    }

    /// <summary>
    ///     Adds shape-level hyperlinks to the hyperlinks list.
    /// </summary>
    /// <param name="presentation">The presentation object.</param>
    /// <param name="autoShape">The AutoShape to check.</param>
    /// <param name="shapeIndex">The shape index.</param>
    /// <param name="hyperlinksList">The list to add hyperlinks to.</param>
    private static void AddShapeLevelHyperlinks(IPresentation presentation, IAutoShape autoShape,
        int shapeIndex, List<object> hyperlinksList)
    {
        if (autoShape.HyperlinkClick != null)
            hyperlinksList.Add(CreateShapeHyperlinkInfo(presentation, autoShape.HyperlinkClick, shapeIndex, "click"));

        if (autoShape.HyperlinkMouseOver != null)
            hyperlinksList.Add(CreateShapeHyperlinkInfo(presentation, autoShape.HyperlinkMouseOver, shapeIndex,
                "mouseover"));
    }

    /// <summary>
    ///     Adds text-level hyperlinks to the hyperlinks list.
    /// </summary>
    /// <param name="presentation">The presentation object.</param>
    /// <param name="autoShape">The AutoShape to check.</param>
    /// <param name="shapeIndex">The shape index.</param>
    /// <param name="hyperlinksList">The list to add hyperlinks to.</param>
    private static void AddTextLevelHyperlinks(IPresentation presentation, IAutoShape autoShape,
        int shapeIndex, List<object> hyperlinksList)
    {
        if (autoShape.TextFrame == null) return;

        foreach (var paragraph in autoShape.TextFrame.Paragraphs)
        foreach (var portion in paragraph.Portions)
        {
            if (portion.PortionFormat.HyperlinkClick != null)
                hyperlinksList.Add(CreateTextHyperlinkInfo(presentation, portion.PortionFormat.HyperlinkClick,
                    shapeIndex, "click", portion.Text));

            if (portion.PortionFormat.HyperlinkMouseOver != null)
                hyperlinksList.Add(CreateTextHyperlinkInfo(presentation, portion.PortionFormat.HyperlinkMouseOver,
                    shapeIndex, "mouseover", portion.Text));
        }
    }

    /// <summary>
    ///     Creates a shape-level hyperlink info object.
    /// </summary>
    /// <param name="presentation">The presentation object.</param>
    /// <param name="hyperlink">The hyperlink.</param>
    /// <param name="shapeIndex">The shape index.</param>
    /// <param name="triggerType">The trigger type (click or mouseover).</param>
    /// <returns>An anonymous object containing hyperlink information.</returns>
    private static object CreateShapeHyperlinkInfo(IPresentation presentation, IHyperlink hyperlink,
        int shapeIndex, string triggerType)
    {
        return new
        {
            shapeIndex,
            level = "shape",
            triggerType,
            url = GetHyperlinkUrl(presentation, hyperlink)
        };
    }

    /// <summary>
    ///     Creates a text-level hyperlink info object.
    /// </summary>
    /// <param name="presentation">The presentation object.</param>
    /// <param name="hyperlink">The hyperlink.</param>
    /// <param name="shapeIndex">The shape index.</param>
    /// <param name="triggerType">The trigger type (click or mouseover).</param>
    /// <param name="text">The text content with the hyperlink.</param>
    /// <returns>An anonymous object containing hyperlink information.</returns>
    private static object CreateTextHyperlinkInfo(IPresentation presentation, IHyperlink hyperlink,
        int shapeIndex, string triggerType, string text)
    {
        return new
        {
            shapeIndex,
            level = "text",
            triggerType,
            text,
            url = GetHyperlinkUrl(presentation, hyperlink)
        };
    }

    /// <summary>
    ///     Gets the URL string from a hyperlink.
    /// </summary>
    /// <param name="presentation">The presentation object.</param>
    /// <param name="hyperlink">The hyperlink.</param>
    /// <returns>The URL string or a description of the internal link.</returns>
    private static string GetHyperlinkUrl(IPresentation presentation, IHyperlink hyperlink)
    {
        if (!string.IsNullOrEmpty(hyperlink.ExternalUrl))
            return hyperlink.ExternalUrl;

        return hyperlink.TargetSlide != null
            ? $"Slide {presentation.Slides.IndexOf(hyperlink.TargetSlide)}"
            : "Internal link";
    }

    /// <summary>
    ///     Creates a hyperlink based on URL or slide target.
    /// </summary>
    /// <param name="presentation">The presentation object.</param>
    /// <param name="url">The external URL (optional).</param>
    /// <param name="slideTargetIndex">The target slide index for internal link (optional).</param>
    /// <returns>A <see cref="HyperlinkCreationResult" /> with the hyperlink and description.</returns>
    /// <exception cref="ArgumentException">Thrown when neither url nor slideTargetIndex is provided.</exception>
    public static HyperlinkCreationResult CreateHyperlink(
        IPresentation presentation, string? url, int? slideTargetIndex)
    {
        if (!string.IsNullOrEmpty(url))
            return new HyperlinkCreationResult(new Hyperlink(url), url);

        if (slideTargetIndex.HasValue)
        {
            if (slideTargetIndex.Value < 0 || slideTargetIndex.Value >= presentation.Slides.Count)
                throw new ArgumentException(
                    $"slideTargetIndex must be between 0 and {presentation.Slides.Count - 1}");

            return new HyperlinkCreationResult(new Hyperlink(presentation.Slides[slideTargetIndex.Value]),
                $"Slide {slideTargetIndex.Value}");
        }

        throw new ArgumentException("Either url or slideTargetIndex must be provided");
    }
}
