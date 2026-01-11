using Aspose.Words;
using Aspose.Words.Notes;
using AsposeMcpServer.Core.Handlers;

namespace AsposeMcpServer.Handlers.Word.Note;

/// <summary>
///     Handler for getting footnotes from Word documents.
/// </summary>
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
    public override string Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var doc = context.Document;
        var notes = WordNoteHelper.GetNotesFromDoc(doc, FootnoteType.Footnote);

        List<object> noteList = [];
        for (var i = 0; i < notes.Count; i++)
        {
            var note = notes[i];
            noteList.Add(new
            {
                noteIndex = i,
                referenceMark = note.ReferenceMark,
                text = note.ToString(SaveFormat.Text).Trim()
            });
        }

        var result = new
        {
            noteType = "footnote",
            count = notes.Count,
            notes = noteList
        };

        return JsonResult(result);
    }
}
