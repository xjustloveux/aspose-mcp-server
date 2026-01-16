using System.Text;
using Aspose.Words;
using Aspose.Words.Notes;
using AsposeMcpServer.Core.Handlers;

namespace AsposeMcpServer.Handlers.Word.Note;

/// <summary>
///     Handler for editing footnotes in Word documents.
/// </summary>
public class EditWordFootnoteHandler : OperationHandlerBase<Document>
{
    /// <inheritdoc />
    public override string Operation => "edit_footnote";

    /// <summary>
    ///     Edits a footnote in the document.
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="parameters">
    ///     Required: text
    ///     Optional: referenceMark, noteIndex (if neither provided, edits first footnote)
    /// </param>
    /// <returns>Success message with edit details.</returns>
    public override string Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var p = ExtractEditFootnoteParameters(parameters);

        var doc = context.Document;
        var notes = WordNoteHelper.GetNotesFromDoc(doc, FootnoteType.Footnote);

        var note = WordNoteHelper.FindNote(notes, p.ReferenceMark, p.NoteIndex);

        if (note == null)
        {
            var availableInfo = notes.Count > 0
                ? $" (document has {notes.Count} footnotes, valid index: 0-{notes.Count - 1})"
                : " (document has no footnotes)";
            throw new ArgumentException(
                $"Specified footnote not found{availableInfo}. Use get_footnotes operation to view available footnotes");
        }

        var oldText = note.ToString(SaveFormat.Text).Trim();
        WordNoteHelper.UpdateNoteText(doc, note, p.NewText);

        MarkModified(context);

        var result = new StringBuilder();
        result.AppendLine("Footnote edited successfully");
        result.AppendLine($"Old text: {oldText}");
        result.AppendLine($"New text: {p.NewText}");
        return result.ToString().TrimEnd();
    }

    /// <summary>
    ///     Extracts edit footnote parameters from operation parameters.
    /// </summary>
    /// <param name="parameters">The operation parameters.</param>
    /// <returns>The extracted edit footnote parameters.</returns>
    private static EditFootnoteParameters ExtractEditFootnoteParameters(OperationParameters parameters)
    {
        return new EditFootnoteParameters(
            parameters.GetRequired<string>("text"),
            parameters.GetOptional<string?>("referenceMark"),
            parameters.GetOptional<int?>("noteIndex")
        );
    }

    /// <summary>
    ///     Record to hold edit footnote parameters.
    /// </summary>
    private record EditFootnoteParameters(string NewText, string? ReferenceMark, int? NoteIndex);
}
