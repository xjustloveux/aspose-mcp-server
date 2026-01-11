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
        var paragraphIndex = parameters.GetOptional("paragraphIndex", 0);
        var borderPosition = parameters.GetOptional<string?>("borderPosition");
        var borderTop = parameters.GetOptional("borderTop", false);
        var borderBottom = parameters.GetOptional("borderBottom", false);
        var borderLeft = parameters.GetOptional("borderLeft", false);
        var borderRight = parameters.GetOptional("borderRight", false);
        var lineStyle = parameters.GetOptional("lineStyle", "single");
        var lineWidth = parameters.GetOptional("lineWidth", 0.5);
        var lineColor = parameters.GetOptional("lineColor", "000000");

        var doc = context.Document;
        var para = WordFormatHelper.GetTargetParagraph(doc, paragraphIndex);
        var borders = para.ParagraphFormat.Borders;

        bool actualBorderTop, actualBorderBottom, actualBorderLeft, actualBorderRight;

        if (!string.IsNullOrEmpty(borderPosition))
        {
            // borderPosition overrides individual flags
            switch (borderPosition.ToLower())
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
                        $"Invalid borderPosition: {borderPosition}. Valid values: all, box, top-bottom, left-right, none");
            }
        }
        else
        {
            // Use individual flags
            actualBorderTop = borderTop;
            actualBorderBottom = borderBottom;
            actualBorderLeft = borderLeft;
            actualBorderRight = borderRight;
        }

        if (actualBorderTop)
        {
            borders.Top.LineStyle = WordFormatHelper.GetLineStyle(lineStyle);
            borders.Top.LineWidth = lineWidth;
            borders.Top.Color = ColorHelper.ParseColor(lineColor);
        }
        else
        {
            borders.Top.LineStyle = LineStyle.None;
        }

        if (actualBorderBottom)
        {
            borders.Bottom.LineStyle = WordFormatHelper.GetLineStyle(lineStyle);
            borders.Bottom.LineWidth = lineWidth;
            borders.Bottom.Color = ColorHelper.ParseColor(lineColor);
        }
        else
        {
            borders.Bottom.LineStyle = LineStyle.None;
        }

        if (actualBorderLeft)
        {
            borders.Left.LineStyle = WordFormatHelper.GetLineStyle(lineStyle);
            borders.Left.LineWidth = lineWidth;
            borders.Left.Color = ColorHelper.ParseColor(lineColor);
        }
        else
        {
            borders.Left.LineStyle = LineStyle.None;
        }

        if (actualBorderRight)
        {
            borders.Right.LineStyle = WordFormatHelper.GetLineStyle(lineStyle);
            borders.Right.LineWidth = lineWidth;
            borders.Right.Color = ColorHelper.ParseColor(lineColor);
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

        return Success($"Paragraph {paragraphIndex} borders set: {bordersDesc}");
    }
}
