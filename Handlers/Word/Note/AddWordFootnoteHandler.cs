using System.Text;
using Aspose.Words;
using Aspose.Words.Notes;
using AsposeMcpServer.Core.Handlers;

namespace AsposeMcpServer.Handlers.Word.Note;

/// <summary>
///     Handler for adding footnotes to Word documents.
/// </summary>
public class AddWordFootnoteHandler : OperationHandlerBase<Document>
{
    /// <inheritdoc />
    public override string Operation => "add_footnote";

    /// <summary>
    ///     Adds a footnote to the document.
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="parameters">
    ///     Required: text
    ///     Optional: paragraphIndex, sectionIndex, referenceText, customMark
    /// </param>
    /// <returns>Success message with footnote details.</returns>
    public override string Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var text = parameters.GetRequired<string>("text");
        var paragraphIndex = parameters.GetOptional<int?>("paragraphIndex");
        var sectionIndex = parameters.GetOptional("sectionIndex", 0);
        var referenceText = parameters.GetOptional<string?>("referenceText");
        var customMark = parameters.GetOptional<string?>("customMark");

        var doc = context.Document;

        if (sectionIndex < 0 || sectionIndex >= doc.Sections.Count)
            throw new ArgumentException($"sectionIndex must be between 0 and {doc.Sections.Count - 1}");

        var builder = new DocumentBuilder(doc);
        Footnote insertedNote;

        if (!string.IsNullOrEmpty(referenceText))
        {
            insertedNote = WordNoteHelper.InsertNoteAtReferenceText(doc, builder, referenceText,
                FootnoteType.Footnote, text, customMark);
        }
        else if (paragraphIndex.HasValue)
        {
            var section = doc.Sections[sectionIndex];
            insertedNote = WordNoteHelper.InsertNoteAtParagraph(builder, section, paragraphIndex.Value,
                FootnoteType.Footnote, text, customMark);
        }
        else
        {
            insertedNote = WordNoteHelper.InsertNoteAtDocumentEnd(builder, FootnoteType.Footnote, text, customMark);
        }

        MarkModified(context);

        var result = new StringBuilder();
        result.AppendLine("Footnote added successfully");
        result.AppendLine($"Text: {text}");
        if (!string.IsNullOrEmpty(insertedNote.ReferenceMark))
            result.AppendLine($"Reference mark: {insertedNote.ReferenceMark}");
        return result.ToString().TrimEnd();
    }
}
