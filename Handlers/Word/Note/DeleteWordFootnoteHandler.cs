using Aspose.Words;
using Aspose.Words.Notes;
using AsposeMcpServer.Core.Handlers;

namespace AsposeMcpServer.Handlers.Word.Note;

/// <summary>
///     Handler for deleting footnotes from Word documents.
/// </summary>
public class DeleteWordFootnoteHandler : OperationHandlerBase<Document>
{
    /// <inheritdoc />
    public override string Operation => "delete_footnote";

    /// <summary>
    ///     Deletes footnotes from the document.
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="parameters">
    ///     Optional: referenceMark, noteIndex (if neither provided, deletes all footnotes)
    /// </param>
    /// <returns>Success message with deletion details.</returns>
    public override string Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var referenceMark = parameters.GetOptional<string?>("referenceMark");
        var noteIndex = parameters.GetOptional<int?>("noteIndex");

        var doc = context.Document;
        var notes = WordNoteHelper.GetNotesFromDoc(doc, FootnoteType.Footnote);

        var deletedCount = 0;

        if (!string.IsNullOrEmpty(referenceMark))
        {
            var note = notes.FirstOrDefault(f => f.ReferenceMark == referenceMark);
            if (note != null)
            {
                note.Remove();
                deletedCount = 1;
            }
        }
        else if (noteIndex.HasValue)
        {
            if (noteIndex.Value >= 0 && noteIndex.Value < notes.Count)
            {
                notes[noteIndex.Value].Remove();
                deletedCount = 1;
            }
            else
            {
                throw new ArgumentException(
                    $"Note index {noteIndex.Value} is out of range (document has {notes.Count} footnotes, valid index: 0-{notes.Count - 1})");
            }
        }
        else
        {
            foreach (var note in notes)
            {
                note.Remove();
                deletedCount++;
            }
        }

        MarkModified(context);

        return Success($"Deleted {deletedCount} footnote(s)");
    }
}
