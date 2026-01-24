using Aspose.Words;
using Aspose.Words.Notes;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Helpers.Word;
using AsposeMcpServer.Results.Word.Note;

namespace AsposeMcpServer.Handlers.Word.Note;

/// <summary>
///     Base handler for getting footnotes/endnotes from Word documents.
/// </summary>
[ResultType(typeof(GetWordNotesResult))]
public abstract class GetWordNotesHandlerBase : OperationHandlerBase<Document>
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
    ///     Gets all notes of the specified type from the document.
    /// </summary>
    /// <param name="context">The document context containing the Word document.</param>
    /// <param name="parameters">The operation parameters (no parameters required).</param>
    /// <returns>
    ///     A <see cref="GetWordNotesResult" /> containing the note type, count,
    ///     and list of <see cref="NoteInfo" /> objects.
    /// </returns>
    public override object Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var doc = context.Document;
        var notes = WordNoteHelper.GetNotesFromDoc(doc, NoteType);

        List<NoteInfo> noteList = [];
        for (var i = 0; i < notes.Count; i++)
        {
            var note = notes[i];
            noteList.Add(new NoteInfo
            {
                NoteIndex = i,
                ReferenceMark = note.ReferenceMark,
                Text = note.ToString(SaveFormat.Text).Trim()
            });
        }

        return new GetWordNotesResult
        {
            NoteType = NoteTypeName.ToLowerInvariant(),
            Count = notes.Count,
            Notes = noteList
        };
    }
}
