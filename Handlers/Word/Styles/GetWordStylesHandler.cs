using Aspose.Words;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Results.Word.Styles;
using WordParagraph = Aspose.Words.Paragraph;
using WordStyle = Aspose.Words.Style;

namespace AsposeMcpServer.Handlers.Word.Styles;

/// <summary>
///     Handler for getting styles from Word documents.
/// </summary>
[ResultType(typeof(GetWordStylesResult))]
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
    public override object
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

        List<StyleInfo> styleList = [];
        foreach (var style in paraStyles)
        {
            var font = style.Font;
            var paraFormat = style.ParagraphFormat;

            var styleInfo = new StyleInfo
            {
                Name = style.Name,
                BuiltIn = style.BuiltIn,
                BasedOn = string.IsNullOrEmpty(style.BaseStyleName) ? null : style.BaseStyleName,
                Font = font.NameAscii == font.NameFarEast ? font.Name : null,
                FontAscii = font.NameAscii != font.NameFarEast ? font.NameAscii : null,
                FontFarEast = font.NameAscii != font.NameFarEast ? font.NameFarEast : null,
                FontSize = font.Size,
                Bold = font.Bold,
                Italic = font.Italic,
                Alignment = paraFormat.Alignment.ToString(),
                SpaceBefore = Math.Abs(paraFormat.SpaceBefore) > FloatTolerance ? paraFormat.SpaceBefore : 0,
                SpaceAfter = Math.Abs(paraFormat.SpaceAfter) > FloatTolerance ? paraFormat.SpaceAfter : 0
            };

            styleList.Add(styleInfo);
        }

        return new GetWordStylesResult
        {
            Count = paraStyles.Count,
            IncludeBuiltIn = p.IncludeBuiltIn,
            Note = p.IncludeBuiltIn
                ? null
                : "Showing custom styles and built-in styles actually used in the document",
            ParagraphStyles = styleList
        };
    }

    private static GetWordStylesParameters ExtractGetWordStylesParameters(OperationParameters parameters)
    {
        return new GetWordStylesParameters(
            parameters.GetOptional("includeBuiltIn", false));
    }

    private sealed record GetWordStylesParameters(bool IncludeBuiltIn);
}
