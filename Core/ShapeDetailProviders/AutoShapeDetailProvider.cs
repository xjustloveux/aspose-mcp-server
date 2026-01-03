using Aspose.Slides;

namespace AsposeMcpServer.Core.ShapeDetailProviders;

/// <summary>
///     Provider for extracting details from AutoShape elements
/// </summary>
public class AutoShapeDetailProvider : IShapeDetailProvider
{
    /// <inheritdoc />
    public string TypeName => "AutoShape";

    /// <inheritdoc />
    public bool CanHandle(IShape shape)
    {
        return shape is IAutoShape;
    }

    /// <inheritdoc />
    public object? GetDetails(IShape shape, IPresentation presentation)
    {
        if (shape is not IAutoShape autoShape)
            return null;

        string? hyperlink = null;
        if (autoShape.HyperlinkClick != null)
            hyperlink = autoShape.HyperlinkClick.ExternalUrl
                        ?? (autoShape.HyperlinkClick.TargetSlide != null
                            ? $"Slide {presentation.Slides.IndexOf(autoShape.HyperlinkClick.TargetSlide)}"
                            : "Internal link");

        var paragraphCount = autoShape.TextFrame?.Paragraphs.Count ?? 0;
        var hasTextFrame = autoShape.TextFrame != null;

        object[]? adjustments = null;
        if (autoShape.Adjustments.Count > 0)
        {
            List<object> adjustmentList = [];
            for (var i = 0; i < autoShape.Adjustments.Count; i++)
                adjustmentList.Add(new { index = i, value = autoShape.Adjustments[i].RawValue });
            adjustments = adjustmentList.ToArray();
        }

        return new
        {
            shapeType = autoShape.ShapeType.ToString(),
            text = autoShape.TextFrame?.Text,
            hasTextFrame,
            paragraphCount,
            hyperlink,
            adjustments
        };
    }
}