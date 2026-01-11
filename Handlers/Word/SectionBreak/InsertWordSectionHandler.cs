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
        var sectionBreakType = parameters.GetRequired<string>("sectionBreakType");
        var insertAtParagraphIndex = parameters.GetOptional<int?>("insertAtParagraphIndex");
        var sectionIndex = parameters.GetOptional<int?>("sectionIndex");

        if (string.IsNullOrEmpty(sectionBreakType))
            throw new ArgumentException("sectionBreakType is required for insert operation");

        var doc = context.Document;
        var builder = new DocumentBuilder(doc);

        var breakType = WordSectionHelper.GetSectionStart(sectionBreakType);

        if (insertAtParagraphIndex.HasValue && insertAtParagraphIndex.Value != -1)
        {
            var actualSectionIndex = sectionIndex ?? 0;
            if (actualSectionIndex < 0 || actualSectionIndex >= doc.Sections.Count)
                throw new ArgumentException(
                    $"sectionIndex must be between 0 and {doc.Sections.Count - 1}, got: {actualSectionIndex}");

            var section = doc.Sections[actualSectionIndex];
            var paragraphs = section.Body.GetChildNodes(NodeType.Paragraph, true).Cast<WordParagraph>().ToList();

            if (paragraphs.Count == 0)
                throw new ArgumentException("Section has no paragraphs to insert section break after");

            if (insertAtParagraphIndex.Value < 0 || insertAtParagraphIndex.Value >= paragraphs.Count)
                throw new ArgumentException(
                    $"insertAtParagraphIndex must be between 0 and {paragraphs.Count - 1}, got: {insertAtParagraphIndex.Value}");

            builder.MoveTo(paragraphs[insertAtParagraphIndex.Value]);
        }
        else
        {
            builder.MoveToDocumentEnd();
        }

        builder.InsertBreak(BreakType.SectionBreakContinuous);
        builder.CurrentSection.PageSetup.SectionStart = breakType;

        MarkModified(context);

        return Success($"Section break inserted ({sectionBreakType})");
    }
}
