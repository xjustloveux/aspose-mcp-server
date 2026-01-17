using Aspose.Words;
using AsposeMcpServer.Core.Handlers;
using WordParagraph = Aspose.Words.Paragraph;
using WordStyle = Aspose.Words.Style;

namespace AsposeMcpServer.Handlers.Word.Styles;

/// <summary>
///     Handler for getting styles from Word documents.
/// </summary>
public class GetWordStylesHandler : OperationHandlerBase<Document>
{
    private const double FloatTolerance = 0.0001;

    /// <inheritdoc />
    public override string Operation => "get_styles";

    /// <summary>
    ///     Gets all styles from the document.
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="parameters">
    ///     Optional: includeBuiltIn
    /// </param>
    /// <returns>JSON string containing style information.</returns>
    public override string
        Execute(OperationContext<Document> context,
            OperationParameters parameters)
    {
        var p = ExtractGetWordStylesParameters(parameters);
        var doc = context.Document;

        List<WordStyle> paraStyles;

        if (p.IncludeBuiltIn)
        {
            paraStyles = doc.Styles
                .Where(s => s.Type == StyleType.Paragraph)
                .OrderBy(s => s.Name)
                .ToList();
        }
        else
        {
            var paragraphs = doc.GetChildNodes(NodeType.Paragraph, true).Cast<WordParagraph>();
            var usedStyleNames = paragraphs
                .Where(para =>
                    para.ParagraphFormat.Style != null && !string.IsNullOrEmpty(para.ParagraphFormat.Style.Name))
                .Select(para => para.ParagraphFormat.Style.Name)
                .ToHashSet();

            paraStyles = doc.Styles
                .Where(s => s.Type == StyleType.Paragraph && (!s.BuiltIn || usedStyleNames.Contains(s.Name)))
                .OrderBy(s => s.Name)
                .ToList();
        }

        List<object> styleList = [];
        foreach (var style in paraStyles)
        {
            var font = style.Font;
            var paraFormat = style.ParagraphFormat;

            var styleInfo = new Dictionary<string, object?>
            {
                ["name"] = style.Name,
                ["builtIn"] = style.BuiltIn
            };

            if (!string.IsNullOrEmpty(style.BaseStyleName))
                styleInfo["basedOn"] = style.BaseStyleName;

            if (font.NameAscii != font.NameFarEast)
            {
                styleInfo["fontAscii"] = font.NameAscii;
                styleInfo["fontFarEast"] = font.NameFarEast;
            }
            else
            {
                styleInfo["font"] = font.Name;
            }

            styleInfo["fontSize"] = font.Size;
            if (font.Bold) styleInfo["bold"] = true;
            if (font.Italic) styleInfo["italic"] = true;

            styleInfo["alignment"] = paraFormat.Alignment.ToString();
            if (Math.Abs(paraFormat.SpaceBefore) > FloatTolerance) styleInfo["spaceBefore"] = paraFormat.SpaceBefore;
            if (Math.Abs(paraFormat.SpaceAfter) > FloatTolerance) styleInfo["spaceAfter"] = paraFormat.SpaceAfter;

            styleList.Add(styleInfo);
        }

        var result = new
        {
            count = paraStyles.Count,
            includeBuiltIn = p.IncludeBuiltIn,
            note = p.IncludeBuiltIn
                ? null
                : "Showing custom styles and built-in styles actually used in the document",
            paragraphStyles = styleList
        };

        return JsonResult(result);
    }

    private static GetWordStylesParameters ExtractGetWordStylesParameters(OperationParameters parameters)
    {
        return new GetWordStylesParameters(
            parameters.GetOptional("includeBuiltIn", false));
    }

    private sealed record GetWordStylesParameters(bool IncludeBuiltIn);
}
