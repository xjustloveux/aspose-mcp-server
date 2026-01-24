using Aspose.Words;
using Aspose.Words.Notes;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Helpers.Word;
using AsposeMcpServer.Results.Common;

namespace AsposeMcpServer.Handlers.Word.Note;

/// <summary>
///     Base handler for deleting footnotes/endnotes from Word documents.
/// </summary>
[ResultType(typeof(SuccessResult))]
public abstract class DeleteWordNoteHandlerBase : OperationHandlerBase<Document>
{
    /// <summary>
    ///     Gets the note type (Footnote or Endnote).
    /// </summary>
    protected abstract FootnoteType NoteType { get; }

    /// <summary>
    ///     Gets the note type name for display (e.g., "Footnote" or "Endnote").
    /// </summary>
    protected abstract string NoteTypeName { get; }

    /// <summary>
    ///     Deletes notes from the document.
    /// </summary>
    /// <param name="context">The document context containing the Word document.</param>
    /// <param name="parameters">
    ///     The operation parameters.
    ///     Optional: referenceMark, noteIndex.
    ///     If neither provided, deletes all notes of this type.
    /// </param>
    /// <returns>A <see cref="SuccessResult" /> with the count of deleted notes.</returns>
    /// <exception cref="ArgumentException">Thrown when noteIndex is out of range.</exception>
    public override object Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var p = ExtractDeleteNoteParameters(parameters);

        var doc = context.Document;
        var notes = WordNoteHelper.GetNotesFromDoc(doc, NoteType);

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
                    $"Note index {p.NoteIndex.Value} is out of range (document has {notes.Count} {NoteTypeName.ToLowerInvariant()}s, valid index: 0-{notes.Count - 1})");
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

        return new SuccessResult { Message = $"Deleted {deletedCount} {NoteTypeName.ToLowerInvariant()}(s)" };
    }

    /// <summary>
    ///     Extracts delete note parameters from operation parameters.
    /// </summary>
    /// <param name="parameters">The operation parameters to extract from.</param>
    /// <returns>The extracted <see cref="DeleteNoteParameters" />.</returns>
    private static DeleteNoteParameters ExtractDeleteNoteParameters(OperationParameters parameters)
    {
        return new DeleteNoteParameters(
            parameters.GetOptional<string?>("referenceMark"),
            parameters.GetOptional<int?>("noteIndex")
        );
    }

    /// <summary>
    ///     Record to hold delete note parameters.
    /// </summary>
    /// <param name="ReferenceMark">Optional reference mark to identify the note.</param>
    /// <param name="NoteIndex">Optional zero-based index of the note.</param>
    private sealed record DeleteNoteParameters(string? ReferenceMark, int? NoteIndex);
}
