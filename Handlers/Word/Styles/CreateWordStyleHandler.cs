using Aspose.Words;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Helpers;

namespace AsposeMcpServer.Handlers.Word.Styles;

/// <summary>
///     Handler for creating styles in Word documents.
/// </summary>
public class CreateWordStyleHandler : OperationHandlerBase<Document>
{
    /// <inheritdoc />
    public override string Operation => "create_style";

    /// <summary>
    ///     Creates a new style in the document.
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="parameters">
    ///     Required: styleName
    ///     Optional: styleType, baseStyle, fontName, fontNameAscii, fontNameFarEast, fontSize,
    ///     bold, italic, underline, color, alignment, spaceBefore, spaceAfter, lineSpacing
    /// </param>
    /// <returns>Success message with style creation details.</returns>
    public override string Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var styleName = parameters.GetRequired<string>("styleName");
        var styleTypeStr = parameters.GetOptional("styleType", "paragraph");
        var baseStyle = parameters.GetOptional<string?>("baseStyle");
        var fontName = parameters.GetOptional<string?>("fontName");
        var fontNameAscii = parameters.GetOptional<string?>("fontNameAscii");
        var fontNameFarEast = parameters.GetOptional<string?>("fontNameFarEast");
        var fontSize = parameters.GetOptional<double?>("fontSize");
        var bold = parameters.GetOptional<bool?>("bold");
        var italic = parameters.GetOptional<bool?>("italic");
        var underline = parameters.GetOptional<bool?>("underline");
        var color = parameters.GetOptional<string?>("color");
        var alignment = parameters.GetOptional<string?>("alignment");
        var spaceBefore = parameters.GetOptional<double?>("spaceBefore");
        var spaceAfter = parameters.GetOptional<double?>("spaceAfter");
        var lineSpacing = parameters.GetOptional<double?>("lineSpacing");

        if (string.IsNullOrEmpty(styleName))
            throw new ArgumentException("styleName is required for create_style operation");

        var doc = context.Document;

        if (doc.Styles[styleName] != null)
            throw new InvalidOperationException($"Style '{styleName}' already exists");

        var styleType = ParseStyleType(styleTypeStr);
        var style = doc.Styles.Add(styleType, styleName);

        SetBaseStyle(doc, style, baseStyle);

        if (styleType != StyleType.List)
            ApplyFontSettings(style, fontName, fontNameAscii, fontNameFarEast, fontSize, bold, italic, underline,
                color);

        if (styleType == StyleType.Paragraph || styleType == StyleType.List)
            ApplyParagraphSettings(style, alignment, spaceBefore, spaceAfter, lineSpacing);

        MarkModified(context);

        return Success($"Style '{styleName}' created successfully");
    }

    private static StyleType ParseStyleType(string styleTypeStr)
    {
        return styleTypeStr.ToLower() switch
        {
            "character" => StyleType.Character,
            "table" => StyleType.Table,
            "list" => StyleType.List,
            _ => StyleType.Paragraph
        };
    }

    private static void SetBaseStyle(Document doc, Style style, string? baseStyle)
    {
        if (string.IsNullOrEmpty(baseStyle)) return;

        var baseStyleObj = doc.Styles[baseStyle];
        if (baseStyleObj != null)
            style.BaseStyleName = baseStyle;
        else
            Console.Error.WriteLine($"[WARN] Base style '{baseStyle}' not found, style will not inherit from it");
    }

    private static void ApplyFontSettings(Style style, string? fontName, string? fontNameAscii, string? fontNameFarEast,
        double? fontSize, bool? bold, bool? italic, bool? underline, string? color)
    {
        if (!string.IsNullOrEmpty(fontNameAscii))
            style.Font.NameAscii = fontNameAscii;

        if (!string.IsNullOrEmpty(fontNameFarEast))
            style.Font.NameFarEast = fontNameFarEast;

        ApplyFontName(style, fontName, fontNameAscii, fontNameFarEast);

        if (fontSize.HasValue)
            style.Font.Size = fontSize.Value;

        if (bold.HasValue)
            style.Font.Bold = bold.Value;

        if (italic.HasValue)
            style.Font.Italic = italic.Value;

        if (underline.HasValue)
            style.Font.Underline = underline.Value ? Underline.Single : Underline.None;

        if (!string.IsNullOrEmpty(color))
            style.Font.Color = ColorHelper.ParseColor(color, true);
    }

    private static void ApplyFontName(Style style, string? fontName, string? fontNameAscii, string? fontNameFarEast)
    {
        if (string.IsNullOrEmpty(fontName)) return;

        if (string.IsNullOrEmpty(fontNameAscii) && string.IsNullOrEmpty(fontNameFarEast))
        {
            style.Font.Name = fontName;
            return;
        }

        if (string.IsNullOrEmpty(fontNameAscii))
            style.Font.NameAscii = fontName;
        if (string.IsNullOrEmpty(fontNameFarEast))
            style.Font.NameFarEast = fontName;
    }

    private static void ApplyParagraphSettings(Style style, string? alignment, double? spaceBefore, double? spaceAfter,
        double? lineSpacing)
    {
        if (!string.IsNullOrEmpty(alignment))
            style.ParagraphFormat.Alignment = alignment.ToLower() switch
            {
                "center" => ParagraphAlignment.Center,
                "right" => ParagraphAlignment.Right,
                "justify" => ParagraphAlignment.Justify,
                _ => ParagraphAlignment.Left
            };

        if (spaceBefore.HasValue)
            style.ParagraphFormat.SpaceBefore = spaceBefore.Value;

        if (spaceAfter.HasValue)
            style.ParagraphFormat.SpaceAfter = spaceAfter.Value;

        if (lineSpacing.HasValue)
        {
            style.ParagraphFormat.LineSpacingRule = LineSpacingRule.Multiple;
            style.ParagraphFormat.LineSpacing = lineSpacing.Value * 12;
        }
    }
}
