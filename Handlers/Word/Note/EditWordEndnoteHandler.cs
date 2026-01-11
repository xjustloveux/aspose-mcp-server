using System.Text;
using Aspose.Words;
using Aspose.Words.Notes;
using AsposeMcpServer.Core.Handlers;

namespace AsposeMcpServer.Handlers.Word.Note;

/// <summary>
///     Handler for editing endnotes in Word documents.
/// </summary>
public class EditWordEndnoteHandler : OperationHandlerBase<Document>
{
    /// <inheritdoc />
    public override string Operation => "edit_endnote";

    /// <summary>
    ///     Edits an endnote in the document.
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="parameters">
    ///     Required: text
    ///     Optional: referenceMark, noteIndex (if neither provided, edits first endnote)
    /// </param>
    /// <returns>Success message with edit details.</returns>
    public override string Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var newText = parameters.GetRequired<string>("text");
        var referenceMark = parameters.GetOptional<string?>("referenceMark");
        var noteIndex = parameters.GetOptional<int?>("noteIndex");

        var doc = context.Document;
        var notes = WordNoteHelper.GetNotesFromDoc(doc, FootnoteType.Endnote);

        var note = WordNoteHelper.FindNote(notes, referenceMark, noteIndex);

        if (note == null)
        {
            var availableInfo = notes.Count > 0
                ? $" (document has {notes.Count} endnotes, valid index: 0-{notes.Count - 1})"
                : " (document has no endnotes)";
            throw new ArgumentException(
                $"Specified endnote not found{availableInfo}. Use get_endnotes operation to view available endnotes");
        }

        var oldText = note.ToString(SaveFormat.Text).Trim();
        WordNoteHelper.UpdateNoteText(doc, note, newText);

        MarkModified(context);

        var result = new StringBuilder();
        result.AppendLine("Endnote edited successfully");
        result.AppendLine($"Old text: {oldText}");
        result.AppendLine($"New text: {newText}");
        return result.ToString().TrimEnd();
    }
}
