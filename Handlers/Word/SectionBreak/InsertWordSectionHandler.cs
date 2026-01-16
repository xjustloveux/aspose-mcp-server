using Aspose.Words;
using AsposeMcpServer.Core.Handlers;
using WordParagraph = Aspose.Words.Paragraph;

namespace AsposeMcpServer.Handlers.Word.SectionBreak;

/// <summary>
///     Handler for inserting section breaks in Word documents.
/// </summary>
public class InsertWordSectionHandler : OperationHandlerBase<Document>
{
    /// <inheritdoc />
    public override string Operation => "insert";

    /// <summary>
    ///     Inserts a section break into the document at specified position.
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="parameters">
    ///     Required: sectionBreakType
    ///     Optional: insertAtParagraphIndex, sectionIndex
    /// </param>
    /// <returns>Success message with section insertion details.</returns>
    public override string Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var p = ExtractInsertWordSectionParameters(parameters);

        if (string.IsNullOrEmpty(p.SectionBreakType))
            throw new ArgumentException("sectionBreakType is required for insert operation");

        var doc = context.Document;
        var builder = new DocumentBuilder(doc);

        var breakType = WordSectionHelper.GetSectionStart(p.SectionBreakType);

        if (p.InsertAtParagraphIndex.HasValue && p.InsertAtParagraphIndex.Value != -1)
        {
            var actualSectionIndex = p.SectionIndex ?? 0;
            if (actualSectionIndex < 0 || actualSectionIndex >= doc.Sections.Count)
                throw new ArgumentException(
                    $"sectionIndex must be between 0 and {doc.Sections.Count - 1}, got: {actualSectionIndex}");

            var section = doc.Sections[actualSectionIndex];
            var paragraphs = section.Body.GetChildNodes(NodeType.Paragraph, true).Cast<WordParagraph>().ToList();

            if (paragraphs.Count == 0)
                throw new ArgumentException("Section has no paragraphs to insert section break after");

            if (p.InsertAtParagraphIndex.Value < 0 || p.InsertAtParagraphIndex.Value >= paragraphs.Count)
                throw new ArgumentException(
                    $"insertAtParagraphIndex must be between 0 and {paragraphs.Count - 1}, got: {p.InsertAtParagraphIndex.Value}");

            builder.MoveTo(paragraphs[p.InsertAtParagraphIndex.Value]);
        }
        else
        {
            builder.MoveToDocumentEnd();
        }

        builder.InsertBreak(BreakType.SectionBreakContinuous);
        builder.CurrentSection.PageSetup.SectionStart = breakType;

        MarkModified(context);

        return Success($"Section break inserted ({p.SectionBreakType})");
    }

    private static InsertWordSectionParameters ExtractInsertWordSectionParameters(OperationParameters parameters)
    {
        return new InsertWordSectionParameters(
            parameters.GetRequired<string>("sectionBreakType"),
            parameters.GetOptional<int?>("insertAtParagraphIndex"),
            parameters.GetOptional<int?>("sectionIndex"));
    }

    private sealed record InsertWordSectionParameters(
        string SectionBreakType,
        int? InsertAtParagraphIndex,
        int? SectionIndex);
}
