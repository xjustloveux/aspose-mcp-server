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
    public override string Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var includeBuiltIn = parameters.GetOptional("includeBuiltIn", false);
        var doc = context.Document;

        List<WordStyle> paraStyles;

        if (includeBuiltIn)
        {
            paraStyles = doc.Styles
                .Where(s => s.Type == StyleType.Paragraph)
                .OrderBy(s => s.Name)
                .ToList();
        }
        else
        {
            var usedStyleNames = new HashSet<string>();
            var paragraphs = doc.GetChildNodes(NodeType.Paragraph, true).Cast<WordParagraph>();
            foreach (var para in paragraphs)
                if (para.ParagraphFormat.Style != null && !string.IsNullOrEmpty(para.ParagraphFormat.Style.Name))
                    usedStyleNames.Add(para.ParagraphFormat.Style.Name);

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
            if (paraFormat.SpaceBefore != 0) styleInfo["spaceBefore"] = paraFormat.SpaceBefore;
            if (paraFormat.SpaceAfter != 0) styleInfo["spaceAfter"] = paraFormat.SpaceAfter;

            styleList.Add(styleInfo);
        }

        var result = new
        {
            count = paraStyles.Count,
            includeBuiltIn,
            note = includeBuiltIn
                ? null
                : "Showing custom styles and built-in styles actually used in the document",
            paragraphStyles = styleList
        };

        return JsonResult(result);
    }
}
