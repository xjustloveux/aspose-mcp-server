using Aspose.Words;
using Aspose.Words.Notes;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Helpers.Word;
using AsposeMcpServer.Results.Word.Note;

namespace AsposeMcpServer.Handlers.Word.Note;

/// <summary>
///     Handler for getting footnotes from Word documents.
/// </summary>
[ResultType(typeof(GetWordNotesResult))]
public class GetWordFootnotesHandler : OperationHandlerBase<Document>
{
    /// <inheritdoc />
    public override string Operation => "get_footnotes";

    /// <summary>
    ///     Gets all footnotes from the document.
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="parameters">No parameters required.</param>
    /// <returns>JSON string containing the list of footnotes.</returns>
    public override object Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var doc = context.Document;
        var notes = WordNoteHelper.GetNotesFromDoc(doc, FootnoteType.Footnote);

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
            NoteType = "footnote",
            Count = notes.Count,
            Notes = noteList
        };
    }
}
