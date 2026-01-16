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
        var p = ExtractApplyWordStyleParameters(parameters);

        if (string.IsNullOrEmpty(p.StyleName))
            throw new ArgumentException("styleName is required for apply_style operation");

        var doc = context.Document;
        var style = doc.Styles[p.StyleName];
        if (style == null)
            throw new ArgumentException($"Style '{p.StyleName}' not found");

        var appliedCount = 0;

        if (p.TableIndex.HasValue)
        {
            var tables = doc.GetChildNodes(NodeType.Table, true).Cast<WordTable>().ToList();
            if (p.TableIndex.Value < 0 || p.TableIndex.Value >= tables.Count)
                throw new ArgumentException($"tableIndex must be between 0 and {tables.Count - 1}");
            tables[p.TableIndex.Value].Style = style;
            appliedCount = 1;
        }
        else if (p.ApplyToAllParagraphs)
        {
            var paragraphs = doc.GetChildNodes(NodeType.Paragraph, true).Cast<WordParagraph>().ToList();
            foreach (var para in paragraphs)
            {
                WordStyleHelper.ApplyStyleToParagraph(para, style, p.StyleName);
                appliedCount++;
            }
        }
        else if (p.ParagraphIndices is { Length: > 0 })
        {
            if (p.SectionIndex < 0 || p.SectionIndex >= doc.Sections.Count)
                throw new ArgumentException($"sectionIndex must be between 0 and {doc.Sections.Count - 1}");

            var section = doc.Sections[p.SectionIndex];
            var paragraphs = section.Body.GetChildNodes(NodeType.Paragraph, true).Cast<WordParagraph>().ToList();

            foreach (var idx in p.ParagraphIndices)
                if (idx >= 0 && idx < paragraphs.Count)
                {
                    WordStyleHelper.ApplyStyleToParagraph(paragraphs[idx], style, p.StyleName);
                    appliedCount++;
                }
        }
        else if (p.ParagraphIndex.HasValue)
        {
            if (p.SectionIndex < 0 || p.SectionIndex >= doc.Sections.Count)
                throw new ArgumentException($"sectionIndex must be between 0 and {doc.Sections.Count - 1}");

            var section = doc.Sections[p.SectionIndex];
            var paragraphs = section.Body.GetChildNodes(NodeType.Paragraph, true).Cast<WordParagraph>().ToList();

            if (p.ParagraphIndex.Value < 0 || p.ParagraphIndex.Value >= paragraphs.Count)
                throw new ArgumentException(
                    $"paragraphIndex must be between 0 and {paragraphs.Count - 1} (section {p.SectionIndex} has {paragraphs.Count} paragraphs, total document paragraphs: {doc.GetChildNodes(NodeType.Paragraph, true).Count})");

            WordStyleHelper.ApplyStyleToParagraph(paragraphs[p.ParagraphIndex.Value], style, p.StyleName);
            appliedCount = 1;
        }
        else
        {
            throw new ArgumentException(
                "Either paragraphIndex, paragraphIndices, tableIndex, or applyToAllParagraphs must be provided");
        }

        MarkModified(context);

        return Success($"Applied style '{p.StyleName}' to {appliedCount} element(s)");
    }

    private static ApplyWordStyleParameters ExtractApplyWordStyleParameters(OperationParameters parameters)
    {
        return new ApplyWordStyleParameters(
            parameters.GetRequired<string>("styleName"),
            parameters.GetOptional<int?>("paragraphIndex"),
            parameters.GetOptional<int[]?>("paragraphIndices"),
            parameters.GetOptional("sectionIndex", 0),
            parameters.GetOptional<int?>("tableIndex"),
            parameters.GetOptional("applyToAllParagraphs", false));
    }

    private sealed record ApplyWordStyleParameters(
        string StyleName,
        int? ParagraphIndex,
        int[]? ParagraphIndices,
        int SectionIndex,
        int? TableIndex,
        bool ApplyToAllParagraphs);
}
