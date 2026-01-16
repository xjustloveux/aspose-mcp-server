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
        var p = ExtractCreateWordStyleParameters(parameters);

        if (string.IsNullOrEmpty(p.StyleName))
            throw new ArgumentException("styleName is required for create_style operation");

        var doc = context.Document;

        if (doc.Styles[p.StyleName] != null)
            throw new InvalidOperationException($"Style '{p.StyleName}' already exists");

        var styleType = ParseStyleType(p.StyleType);
        var style = doc.Styles.Add(styleType, p.StyleName);

        SetBaseStyle(doc, style, p.BaseStyle);

        if (styleType != StyleType.List)
            ApplyFontSettings(style, p.FontName, p.FontNameAscii, p.FontNameFarEast, p.FontSize, p.Bold, p.Italic,
                p.Underline,
                p.Color);

        if (styleType == StyleType.Paragraph || styleType == StyleType.List)
            ApplyParagraphSettings(style, p.Alignment, p.SpaceBefore, p.SpaceAfter, p.LineSpacing);

        MarkModified(context);

        return Success($"Style '{p.StyleName}' created successfully");
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

    private static CreateWordStyleParameters ExtractCreateWordStyleParameters(OperationParameters parameters)
    {
        return new CreateWordStyleParameters(
            parameters.GetRequired<string>("styleName"),
            parameters.GetOptional("styleType", "paragraph"),
            parameters.GetOptional<string?>("baseStyle"),
            parameters.GetOptional<string?>("fontName"),
            parameters.GetOptional<string?>("fontNameAscii"),
            parameters.GetOptional<string?>("fontNameFarEast"),
            parameters.GetOptional<double?>("fontSize"),
            parameters.GetOptional<bool?>("bold"),
            parameters.GetOptional<bool?>("italic"),
            parameters.GetOptional<bool?>("underline"),
            parameters.GetOptional<string?>("color"),
            parameters.GetOptional<string?>("alignment"),
            parameters.GetOptional<double?>("spaceBefore"),
            parameters.GetOptional<double?>("spaceAfter"),
            parameters.GetOptional<double?>("lineSpacing"));
    }

    private record CreateWordStyleParameters(
        string StyleName,
        string StyleType,
        string? BaseStyle,
        string? FontName,
        string? FontNameAscii,
        string? FontNameFarEast,
        double? FontSize,
        bool? Bold,
        bool? Italic,
        bool? Underline,
        string? Color,
        string? Alignment,
        double? SpaceBefore,
        double? SpaceAfter,
        double? LineSpacing);
}
