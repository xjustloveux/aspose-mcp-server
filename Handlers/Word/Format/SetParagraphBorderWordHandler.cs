using Aspose.Words;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Helpers;

namespace AsposeMcpServer.Handlers.Word.Format;

/// <summary>
///     Handler for setting paragraph border in Word documents.
/// </summary>
public class SetParagraphBorderWordHandler : OperationHandlerBase<Document>
{
    /// <inheritdoc />
    public override string Operation => "set_paragraph_border";

    /// <summary>
    ///     Sets paragraph border properties.
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="parameters">
    ///     Required: paragraphIndex
    ///     Optional: borderPosition, borderTop, borderBottom, borderLeft, borderRight, lineStyle, lineWidth, lineColor
    /// </param>
    /// <returns>Success message.</returns>
    public override string Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var p = ExtractSetParagraphBorderParameters(parameters);

        var doc = context.Document;
        var para = WordFormatHelper.GetTargetParagraph(doc, p.ParagraphIndex);
        var borders = para.ParagraphFormat.Borders;

        bool actualBorderTop, actualBorderBottom, actualBorderLeft, actualBorderRight;

        if (!string.IsNullOrEmpty(p.BorderPosition))
        {
            switch (p.BorderPosition.ToLower())
            {
                case "all":
                case "box":
                    actualBorderTop = actualBorderBottom = actualBorderLeft = actualBorderRight = true;
                    break;
                case "top-bottom":
                    actualBorderTop = actualBorderBottom = true;
                    actualBorderLeft = actualBorderRight = false;
                    break;
                case "left-right":
                    actualBorderTop = actualBorderBottom = false;
                    actualBorderLeft = actualBorderRight = true;
                    break;
                case "none":
                    actualBorderTop = actualBorderBottom = actualBorderLeft = actualBorderRight = false;
                    break;
                default:
                    throw new ArgumentException(
                        $"Invalid borderPosition: {p.BorderPosition}. Valid values: all, box, top-bottom, left-right, none");
            }
        }
        else
        {
            actualBorderTop = p.BorderTop;
            actualBorderBottom = p.BorderBottom;
            actualBorderLeft = p.BorderLeft;
            actualBorderRight = p.BorderRight;
        }

        if (actualBorderTop)
        {
            borders.Top.LineStyle = WordFormatHelper.GetLineStyle(p.LineStyle);
            borders.Top.LineWidth = p.LineWidth;
            borders.Top.Color = ColorHelper.ParseColor(p.LineColor);
        }
        else
        {
            borders.Top.LineStyle = LineStyle.None;
        }

        if (actualBorderBottom)
        {
            borders.Bottom.LineStyle = WordFormatHelper.GetLineStyle(p.LineStyle);
            borders.Bottom.LineWidth = p.LineWidth;
            borders.Bottom.Color = ColorHelper.ParseColor(p.LineColor);
        }
        else
        {
            borders.Bottom.LineStyle = LineStyle.None;
        }

        if (actualBorderLeft)
        {
            borders.Left.LineStyle = WordFormatHelper.GetLineStyle(p.LineStyle);
            borders.Left.LineWidth = p.LineWidth;
            borders.Left.Color = ColorHelper.ParseColor(p.LineColor);
        }
        else
        {
            borders.Left.LineStyle = LineStyle.None;
        }

        if (actualBorderRight)
        {
            borders.Right.LineStyle = WordFormatHelper.GetLineStyle(p.LineStyle);
            borders.Right.LineWidth = p.LineWidth;
            borders.Right.Color = ColorHelper.ParseColor(p.LineColor);
        }
        else
        {
            borders.Right.LineStyle = LineStyle.None;
        }

        MarkModified(context);

        List<string> enabledBorders = [];
        if (actualBorderTop) enabledBorders.Add("Top");
        if (actualBorderBottom) enabledBorders.Add("Bottom");
        if (actualBorderLeft) enabledBorders.Add("Left");
        if (actualBorderRight) enabledBorders.Add("Right");

        var bordersDesc = enabledBorders.Count > 0 ? string.Join(", ", enabledBorders) : "None";

        return Success($"Paragraph {p.ParagraphIndex} borders set: {bordersDesc}");
    }

    private static SetParagraphBorderParameters ExtractSetParagraphBorderParameters(OperationParameters parameters)
    {
        return new SetParagraphBorderParameters(
            parameters.GetOptional("paragraphIndex", 0),
            parameters.GetOptional<string?>("borderPosition"),
            parameters.GetOptional("borderTop", false),
            parameters.GetOptional("borderBottom", false),
            parameters.GetOptional("borderLeft", false),
            parameters.GetOptional("borderRight", false),
            parameters.GetOptional("lineStyle", "single"),
            parameters.GetOptional("lineWidth", 0.5),
            parameters.GetOptional("lineColor", "000000"));
    }

    private record SetParagraphBorderParameters(
        int ParagraphIndex,
        string? BorderPosition,
        bool BorderTop,
        bool BorderBottom,
        bool BorderLeft,
        bool BorderRight,
        string LineStyle,
        double LineWidth,
        string LineColor);
}
