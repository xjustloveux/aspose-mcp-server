using Aspose.Words;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Helpers;
using AsposeMcpServer.Results.Common;

namespace AsposeMcpServer.Handlers.Word.Styles;

/// <summary>
///     Handler for creating styles in Word documents.
/// </summary>
[ResultType(typeof(SuccessResult))]
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
    public override object Execute(OperationContext<Document> context, OperationParameters parameters)
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
            ApplyFontSettings(style, p);

        if (styleType == StyleType.Paragraph || styleType == StyleType.List)
            ApplyParagraphSettings(style, p.Alignment, p.SpaceBefore, p.SpaceAfter, p.LineSpacing);

        MarkModified(context);

        return new SuccessResult { Message = $"Style '{p.StyleName}' created successfully" };
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

    /// <summary>
    ///     Applies font settings to the style.
    /// </summary>
    /// <param name="style">The style to apply font settings to.</param>
    /// <param name="p">The parameters containing font settings.</param>
    private static void ApplyFontSettings(Style style, CreateWordStyleParameters p)
    {
        if (!string.IsNullOrEmpty(p.FontNameAscii))
            style.Font.NameAscii = p.FontNameAscii;

        if (!string.IsNullOrEmpty(p.FontNameFarEast))
            style.Font.NameFarEast = p.FontNameFarEast;

        ApplyFontName(style, p.FontName, p.FontNameAscii, p.FontNameFarEast);

        if (p.FontSize.HasValue)
            style.Font.Size = p.FontSize.Value;

        if (p.Bold.HasValue)
            style.Font.Bold = p.Bold.Value;

        if (p.Italic.HasValue)
            style.Font.Italic = p.Italic.Value;

        if (p.Underline.HasValue)
            style.Font.Underline = p.Underline.Value ? Underline.Single : Underline.None;

        if (!string.IsNullOrEmpty(p.Color))
            style.Font.Color = ColorHelper.ParseColor(p.Color, true);
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

    /// <summary>
    ///     Record to hold create word style parameters.
    /// </summary>
    /// <param name="StyleName">The style name.</param>
    /// <param name="StyleType">The style type.</param>
    /// <param name="BaseStyle">The base style name.</param>
    /// <param name="FontName">The font name.</param>
    /// <param name="FontNameAscii">The ASCII font name.</param>
    /// <param name="FontNameFarEast">The Far East font name.</param>
    /// <param name="FontSize">The font size.</param>
    /// <param name="Bold">Whether to apply bold.</param>
    /// <param name="Italic">Whether to apply italic.</param>
    /// <param name="Underline">Whether to apply underline.</param>
    /// <param name="Color">The font color.</param>
    /// <param name="Alignment">The paragraph alignment.</param>
    /// <param name="SpaceBefore">The space before in points.</param>
    /// <param name="SpaceAfter">The space after in points.</param>
    /// <param name="LineSpacing">The line spacing multiplier.</param>
    private sealed record CreateWordStyleParameters(
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
