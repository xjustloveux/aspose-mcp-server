using System.Text;
using Aspose.Words;
using Aspose.Words.Notes;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Helpers.Word;
using AsposeMcpServer.Results.Common;

namespace AsposeMcpServer.Handlers.Word.Note;

/// <summary>
///     Handler for editing endnotes in Word documents.
/// </summary>
[ResultType(typeof(SuccessResult))]
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
    public override object Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var p = ExtractEditEndnoteParameters(parameters);

        var doc = context.Document;
        var notes = WordNoteHelper.GetNotesFromDoc(doc, FootnoteType.Endnote);

        var note = WordNoteHelper.FindNote(notes, p.ReferenceMark, p.NoteIndex);

        if (note == null)
        {
            var availableInfo = notes.Count > 0
                ? $" (document has {notes.Count} endnotes, valid index: 0-{notes.Count - 1})"
                : " (document has no endnotes)";
            throw new ArgumentException(
                $"Specified endnote not found{availableInfo}. Use get_endnotes operation to view available endnotes");
        }

        var oldText = note.ToString(SaveFormat.Text).Trim();
        WordNoteHelper.UpdateNoteText(doc, note, p.NewText);

        MarkModified(context);

        return BuildSuccessResult(oldText, p.NewText);
    }

    /// <summary>
    ///     Builds the success result for editing an endnote.
    /// </summary>
    /// <param name="oldText">The old endnote text.</param>
    /// <param name="newText">The new endnote text.</param>
    /// <returns>The success result.</returns>
    private static SuccessResult BuildSuccessResult(string oldText, string newText)
    {
        var result = new StringBuilder();
        result.AppendLine("Endnote edited successfully");
        result.AppendLine($"Old text: {oldText}");
        result.AppendLine($"New text: {newText}");
        return new SuccessResult { Message = result.ToString().TrimEnd() };
    }

    /// <summary>
    ///     Extracts edit endnote parameters from operation parameters.
    /// </summary>
    /// <param name="parameters">The operation parameters.</param>
    /// <returns>The extracted edit endnote parameters.</returns>
    private static EditEndnoteParameters ExtractEditEndnoteParameters(OperationParameters parameters)
    {
        return new EditEndnoteParameters(
            parameters.GetRequired<string>("text"),
            parameters.GetOptional<string?>("referenceMark"),
            parameters.GetOptional<int?>("noteIndex")
        );
    }

    /// <summary>
    ///     Record to hold edit endnote parameters.
    /// </summary>
    private sealed record EditEndnoteParameters(string NewText, string? ReferenceMark, int? NoteIndex);
}
