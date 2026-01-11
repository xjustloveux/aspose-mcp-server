using Aspose.Words;
using AsposeMcpServer.Core.Handlers;
using WordParagraph = Aspose.Words.Paragraph;
using WordTable = Aspose.Words.Tables.Table;

namespace AsposeMcpServer.Handlers.Word.Styles;

/// <summary>
///     Handler for applying styles to paragraphs or tables in Word documents.
/// </summary>
public class ApplyWordStyleHandler : OperationHandlerBase<Document>
{
    /// <inheritdoc />
    public override string Operation => "apply_style";

    /// <summary>
    ///     Applies a style to paragraphs or tables.
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="parameters">
    ///     Required: styleName
    ///     Optional: paragraphIndex, paragraphIndices, sectionIndex, tableIndex, applyToAllParagraphs
    /// </param>
    /// <returns>Success message with application details.</returns>
    public override string Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var styleName = parameters.GetRequired<string>("styleName");
        var paragraphIndex = parameters.GetOptional<int?>("paragraphIndex");
        var paragraphIndices = parameters.GetOptional<int[]?>("paragraphIndices");
        var sectionIndex = parameters.GetOptional("sectionIndex", 0);
        var tableIndex = parameters.GetOptional<int?>("tableIndex");
        var applyToAllParagraphs = parameters.GetOptional("applyToAllParagraphs", false);

        if (string.IsNullOrEmpty(styleName))
            throw new ArgumentException("styleName is required for apply_style operation");

        var doc = context.Document;
        var style = doc.Styles[styleName];
        if (style == null)
            throw new ArgumentException($"Style '{styleName}' not found");

        var appliedCount = 0;

        if (tableIndex.HasValue)
        {
            var tables = doc.GetChildNodes(NodeType.Table, true).Cast<WordTable>().ToList();
            if (tableIndex.Value < 0 || tableIndex.Value >= tables.Count)
                throw new ArgumentException($"tableIndex must be between 0 and {tables.Count - 1}");
            tables[tableIndex.Value].Style = style;
            appliedCount = 1;
        }
        else if (applyToAllParagraphs)
        {
            var paragraphs = doc.GetChildNodes(NodeType.Paragraph, true).Cast<WordParagraph>().ToList();
            foreach (var para in paragraphs)
            {
                WordStyleHelper.ApplyStyleToParagraph(para, style, styleName);
                appliedCount++;
            }
        }
        else if (paragraphIndices is { Length: > 0 })
        {
            if (sectionIndex < 0 || sectionIndex >= doc.Sections.Count)
                throw new ArgumentException($"sectionIndex must be between 0 and {doc.Sections.Count - 1}");

            var section = doc.Sections[sectionIndex];
            var paragraphs = section.Body.GetChildNodes(NodeType.Paragraph, true).Cast<WordParagraph>().ToList();

            foreach (var idx in paragraphIndices)
                if (idx >= 0 && idx < paragraphs.Count)
                {
                    WordStyleHelper.ApplyStyleToParagraph(paragraphs[idx], style, styleName);
                    appliedCount++;
                }
        }
        else if (paragraphIndex.HasValue)
        {
            if (sectionIndex < 0 || sectionIndex >= doc.Sections.Count)
                throw new ArgumentException($"sectionIndex must be between 0 and {doc.Sections.Count - 1}");

            var section = doc.Sections[sectionIndex];
            var paragraphs = section.Body.GetChildNodes(NodeType.Paragraph, true).Cast<WordParagraph>().ToList();

            if (paragraphIndex.Value < 0 || paragraphIndex.Value >= paragraphs.Count)
                throw new ArgumentException(
                    $"paragraphIndex must be between 0 and {paragraphs.Count - 1} (section {sectionIndex} has {paragraphs.Count} paragraphs, total document paragraphs: {doc.GetChildNodes(NodeType.Paragraph, true).Count})");

            WordStyleHelper.ApplyStyleToParagraph(paragraphs[paragraphIndex.Value], style, styleName);
            appliedCount = 1;
        }
        else
        {
            throw new ArgumentException(
                "Either paragraphIndex, paragraphIndices, tableIndex, or applyToAllParagraphs must be provided");
        }

        MarkModified(context);

        return Success($"Applied style '{styleName}' to {appliedCount} element(s)");
    }
}
