using Aspose.Words;
using Aspose.Words.Notes;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Helpers.Word;
using AsposeMcpServer.Results.Common;

namespace AsposeMcpServer.Handlers.Word.Note;

/// <summary>
///     Handler for deleting endnotes from Word documents.
/// </summary>
[ResultType(typeof(SuccessResult))]
public class DeleteWordEndnoteHandler : OperationHandlerBase<Document>
{
    /// <inheritdoc />
    public override string Operation => "delete_endnote";

    /// <summary>
    ///     Deletes endnotes from the document.
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="parameters">
    ///     Optional: referenceMark, noteIndex (if neither provided, deletes all endnotes)
    /// </param>
    /// <returns>Success message with deletion details.</returns>
    public override object Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var p = ExtractDeleteEndnoteParameters(parameters);

        var doc = context.Document;
        var notes = WordNoteHelper.GetNotesFromDoc(doc, FootnoteType.Endnote);

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
                    $"Note index {p.NoteIndex.Value} is out of range (document has {notes.Count} endnotes, valid index: 0-{notes.Count - 1})");
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

        return new SuccessResult { Message = $"Deleted {deletedCount} endnote(s)" };
    }

    /// <summary>
    ///     Extracts delete endnote parameters from operation parameters.
    /// </summary>
    /// <param name="parameters">The operation parameters.</param>
    /// <returns>The extracted delete endnote parameters.</returns>
    private static DeleteEndnoteParameters ExtractDeleteEndnoteParameters(OperationParameters parameters)
    {
        return new DeleteEndnoteParameters(
            parameters.GetOptional<string?>("referenceMark"),
            parameters.GetOptional<int?>("noteIndex")
        );
    }

    /// <summary>
    ///     Record to hold delete endnote parameters.
    /// </summary>
    private sealed record DeleteEndnoteParameters(string? ReferenceMark, int? NoteIndex);
}
