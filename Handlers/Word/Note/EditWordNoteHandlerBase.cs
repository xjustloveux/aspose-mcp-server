using System.Text;
using Aspose.Words;
using Aspose.Words.Notes;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Helpers.Word;
using AsposeMcpServer.Results.Common;

namespace AsposeMcpServer.Handlers.Word.Note;

/// <summary>
///     Base handler for editing footnotes/endnotes in Word documents.
/// </summary>
[ResultType(typeof(SuccessResult))]
public abstract class EditWordNoteHandlerBase : OperationHandlerBase<Document>
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
    ///     Edits a note in the document.
    /// </summary>
    /// <param name="context">The document context containing the Word document.</param>
    /// <param name="parameters">
    ///     The operation parameters.
    ///     Required: text (the new note content).
    ///     Optional: referenceMark, noteIndex.
    ///     If neither identifier provided, edits the first note.
    /// </param>
    /// <returns>A <see cref="SuccessResult" /> with old and new text details.</returns>
    /// <exception cref="ArgumentException">Thrown when the specified note is not found.</exception>
    public override object Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var p = ExtractEditNoteParameters(parameters);

        var doc = context.Document;
        var notes = WordNoteHelper.GetNotesFromDoc(doc, NoteType);

        var note = WordNoteHelper.FindNote(notes, p.ReferenceMark, p.NoteIndex);

        if (note == null)
        {
            var availableInfo = notes.Count > 0
                ? $" (document has {notes.Count} {NoteTypeName.ToLowerInvariant()}s, valid index: 0-{notes.Count - 1})"
                : $" (document has no {NoteTypeName.ToLowerInvariant()}s)";
            throw new ArgumentException(
                $"Specified {NoteTypeName.ToLowerInvariant()} not found{availableInfo}. Use get_{NoteTypeName.ToLowerInvariant()}s operation to view available {NoteTypeName.ToLowerInvariant()}s");
        }

        var oldText = note.ToString(SaveFormat.Text).Trim();
        WordNoteHelper.UpdateNoteText(doc, note, p.NewText);

        MarkModified(context);

        return BuildSuccessResult(oldText, p.NewText);
    }

    /// <summary>
    ///     Builds the success result for editing a note.
    /// </summary>
    /// <param name="oldText">The original note text.</param>
    /// <param name="newText">The new note text.</param>
    /// <returns>A <see cref="SuccessResult" /> containing the operation details.</returns>
    private SuccessResult BuildSuccessResult(string oldText, string newText)
    {
        var result = new StringBuilder();
        result.AppendLine($"{NoteTypeName} edited successfully");
        result.AppendLine($"Old text: {oldText}");
        result.AppendLine($"New text: {newText}");
        return new SuccessResult { Message = result.ToString().TrimEnd() };
    }

    /// <summary>
    ///     Extracts edit note parameters from operation parameters.
    /// </summary>
    /// <param name="parameters">The operation parameters to extract from.</param>
    /// <returns>The extracted <see cref="EditNoteParameters" />.</returns>
    private static EditNoteParameters ExtractEditNoteParameters(OperationParameters parameters)
    {
        return new EditNoteParameters(
            parameters.GetRequired<string>("text"),
            parameters.GetOptional<string?>("referenceMark"),
            parameters.GetOptional<int?>("noteIndex")
        );
    }

    /// <summary>
    ///     Record to hold edit note parameters.
    /// </summary>
    /// <param name="NewText">The new text content for the note.</param>
    /// <param name="ReferenceMark">Optional reference mark to identify the note.</param>
    /// <param name="NoteIndex">Optional zero-based index of the note.</param>
    private sealed record EditNoteParameters(string NewText, string? ReferenceMark, int? NoteIndex);
}
