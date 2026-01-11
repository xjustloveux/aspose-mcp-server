using Aspose.Slides;

namespace AsposeMcpServer.Handlers.PowerPoint.Hyperlink;

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

            if (autoShape.HyperlinkClick != null)
            {
                var targetSlide = autoShape.HyperlinkClick.TargetSlide;
                var url = autoShape.HyperlinkClick.ExternalUrl
                          ?? (targetSlide != null
                              ? $"Slide {presentation.Slides.IndexOf(targetSlide)}"
                              : "Internal link");

                hyperlinksList.Add(new
                {
                    shapeIndex,
                    level = "shape",
                    triggerType = "click",
                    url
                });
            }

            if (autoShape.HyperlinkMouseOver != null)
            {
                var targetSlide = autoShape.HyperlinkMouseOver.TargetSlide;
                var url = autoShape.HyperlinkMouseOver.ExternalUrl
                          ?? (targetSlide != null
                              ? $"Slide {presentation.Slides.IndexOf(targetSlide)}"
                              : "Internal link");

                hyperlinksList.Add(new
                {
                    shapeIndex,
                    level = "shape",
                    triggerType = "mouseover",
                    url
                });
            }

            if (autoShape.TextFrame == null) continue;

            foreach (var paragraph in autoShape.TextFrame.Paragraphs)
            foreach (var portion in paragraph.Portions)
            {
                if (portion.PortionFormat.HyperlinkClick != null)
                {
                    var targetSlide = portion.PortionFormat.HyperlinkClick.TargetSlide;
                    var url = portion.PortionFormat.HyperlinkClick.ExternalUrl
                              ?? (targetSlide != null
                                  ? $"Slide {presentation.Slides.IndexOf(targetSlide)}"
                                  : "Internal link");

                    hyperlinksList.Add(new
                    {
                        shapeIndex,
                        level = "text",
                        triggerType = "click",
                        text = portion.Text,
                        url
                    });
                }

                if (portion.PortionFormat.HyperlinkMouseOver != null)
                {
                    var targetSlide = portion.PortionFormat.HyperlinkMouseOver.TargetSlide;
                    var url = portion.PortionFormat.HyperlinkMouseOver.ExternalUrl
                              ?? (targetSlide != null
                                  ? $"Slide {presentation.Slides.IndexOf(targetSlide)}"
                                  : "Internal link");

                    hyperlinksList.Add(new
                    {
                        shapeIndex,
                        level = "text",
                        triggerType = "mouseover",
                        text = portion.Text,
                        url
                    });
                }
            }
        }

        return hyperlinksList;
    }

    /// <summary>
    ///     Creates a hyperlink based on URL or slide target.
    /// </summary>
    /// <param name="presentation">The presentation object.</param>
    /// <param name="url">The external URL (optional).</param>
    /// <param name="slideTargetIndex">The target slide index for internal link (optional).</param>
    /// <returns>A tuple containing the hyperlink and its description.</returns>
    /// <exception cref="ArgumentException">Thrown when neither url nor slideTargetIndex is provided.</exception>
    public static (IHyperlink hyperlink, string description) CreateHyperlink(
        IPresentation presentation, string? url, int? slideTargetIndex)
    {
        if (!string.IsNullOrEmpty(url)) return (new Aspose.Slides.Hyperlink(url), url);

        if (slideTargetIndex.HasValue)
        {
            if (slideTargetIndex.Value < 0 || slideTargetIndex.Value >= presentation.Slides.Count)
                throw new ArgumentException(
                    $"slideTargetIndex must be between 0 and {presentation.Slides.Count - 1}");

            return (new Aspose.Slides.Hyperlink(presentation.Slides[slideTargetIndex.Value]),
                $"Slide {slideTargetIndex.Value}");
        }

        throw new ArgumentException("Either url or slideTargetIndex must be provided");
    }
}
