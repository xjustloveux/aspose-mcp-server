using System.Drawing;
using Aspose.Slides;
using AsposeMcpServer.Core.ShapeDetailProviders.Details;

namespace AsposeMcpServer.Core.ShapeDetailProviders.Providers;

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
    public ShapeDetails? GetDetails(IShape shape, IPresentation presentation)
    {
        if (shape is not IAutoShape autoShape)
            return null;

        var hyperlink = GetHyperlinkText(autoShape.HyperlinkClick, presentation);

        var paragraphCount = autoShape.TextFrame?.Paragraphs.Count ?? 0;
        var hasTextFrame = autoShape.TextFrame != null;

        IReadOnlyList<AdjustmentInfo>? adjustments = null;
        if (autoShape.Adjustments.Count > 0)
        {
            List<AdjustmentInfo> adjustmentList = [];
            for (var i = 0; i < autoShape.Adjustments.Count; i++)
                adjustmentList.Add(new AdjustmentInfo
                {
                    Index = i,
                    Value = autoShape.Adjustments[i].RawValue
                });
            adjustments = adjustmentList;
        }

        string? fillColor = null;
        float? transparency = null;
        if (autoShape.FillFormat is { FillType: FillType.Solid })
        {
            var color = autoShape.FillFormat.SolidFillColor.Color;
            if (color != Color.Empty)
            {
                fillColor = $"#{color.R:X2}{color.G:X2}{color.B:X2}";
                transparency = color.A < 255 ? 1f - color.A / 255f : null;
            }
        }

        string? lineColor = null;
        double? lineWidth = null;
        string? lineDashStyle = null;
        var lineFormat = autoShape.LineFormat;
        if (lineFormat != null)
        {
            if (lineFormat.FillFormat is { FillType: FillType.Solid })
            {
                var lc = lineFormat.FillFormat.SolidFillColor.Color;
                if (lc != Color.Empty)
                    lineColor = $"#{lc.R:X2}{lc.G:X2}{lc.B:X2}";
            }

            if (lineFormat.Width is > 0 and not double.NaN)
                lineWidth = lineFormat.Width;

            lineDashStyle = lineFormat.DashStyle.ToString();
            if (lineDashStyle == "NotDefined")
                lineDashStyle = null;
        }

        return new AutoShapeDetails
        {
            ShapeType = autoShape.ShapeType.ToString(),
            Text = autoShape.TextFrame?.Text,
            HasTextFrame = hasTextFrame,
            ParagraphCount = paragraphCount,
            Hyperlink = hyperlink,
            FillType = autoShape.FillFormat?.FillType.ToString(),
            FillColor = fillColor,
            Transparency = transparency,
            LineColor = lineColor,
            LineWidth = lineWidth,
            LineDashStyle = lineDashStyle,
            Adjustments = adjustments
        };
    }

    /// <summary>
    ///     Gets the hyperlink text from a hyperlink click action.
    /// </summary>
    /// <param name="hyperlink">The hyperlink to extract text from.</param>
    /// <param name="presentation">The presentation for resolving slide references.</param>
    /// <returns>The hyperlink text, or null if no hyperlink is set.</returns>
    internal static string? GetHyperlinkText(IHyperlink? hyperlink, IPresentation presentation)
    {
        if (hyperlink == null)
            return null;

        return hyperlink.ExternalUrl
               ?? (hyperlink.TargetSlide != null
                   ? $"Slide {presentation.Slides.IndexOf(hyperlink.TargetSlide)}"
                   : "Internal link");
    }
}
