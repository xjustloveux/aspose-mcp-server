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
        var p = ExtractDeleteFootnoteParameters(parameters);

        var doc = context.Document;
        var notes = WordNoteHelper.GetNotesFromDoc(doc, FootnoteType.Footnote);

        var deletedCount = 0;

        if (!string.IsNullOrEmpty(p.ReferenceMark))
        {
            var note = notes.FirstOrDefault(f => f.ReferenceMark == p.ReferenceMark);
            if (note != null)
            {
                note.Remove();
                deletedCount = 1;
            }
        }
        else if (p.NoteIndex.HasValue)
        {
            if (p.NoteIndex.Value >= 0 && p.NoteIndex.Value < notes.Count)
            {
                notes[p.NoteIndex.Value].Remove();
                deletedCount = 1;
            }
            else
            {
                throw new ArgumentException(
                    $"Note index {p.NoteIndex.Value} is out of range (document has {notes.Count} footnotes, valid index: 0-{notes.Count - 1})");
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

    /// <summary>
    ///     Extracts delete footnote parameters from operation parameters.
    /// </summary>
    /// <param name="parameters">The operation parameters.</param>
    /// <returns>The extracted delete footnote parameters.</returns>
    private static DeleteFootnoteParameters ExtractDeleteFootnoteParameters(OperationParameters parameters)
    {
        return new DeleteFootnoteParameters(
            parameters.GetOptional<string?>("referenceMark"),
            parameters.GetOptional<int?>("noteIndex")
        );
    }

    /// <summary>
    ///     Record to hold delete footnote parameters.
    /// </summary>
    private record DeleteFootnoteParameters(string? ReferenceMark, int? NoteIndex);
}
