using System.Text;
using Aspose.Words;
using Aspose.Words.Notes;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Helpers.Word;
using AsposeMcpServer.Results.Common;

namespace AsposeMcpServer.Handlers.Word.Note;

/// <summary>
///     Base handler for adding footnotes/endnotes to Word documents.
/// </summary>
[ResultType(typeof(SuccessResult))]
public abstract class AddWordNoteHandlerBase : OperationHandlerBase<Document>
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
    ///     Adds a note to the document.
    /// </summary>
    /// <param name="context">The document context containing the Word document.</param>
    /// <param name="parameters">
    ///     The operation parameters.
    ///     Required: text (the note content).
    ///     Optional: paragraphIndex, sectionIndex, referenceText, customMark.
    /// </param>
    /// <returns>A <see cref="SuccessResult" /> with note details including reference mark.</returns>
    /// <exception cref="ArgumentException">Thrown when sectionIndex is out of range.</exception>
    public override object Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var p = ExtractAddNoteParameters(parameters);

        var doc = context.Document;

        if (p.SectionIndex < 0 || p.SectionIndex >= doc.Sections.Count)
            throw new ArgumentException($"sectionIndex must be between 0 and {doc.Sections.Count - 1}");

        var builder = new DocumentBuilder(doc);
        Footnote insertedNote;

        if (!string.IsNullOrEmpty(p.ReferenceText))
        {
            insertedNote = WordNoteHelper.InsertNoteAtReferenceText(doc, builder, p.ReferenceText,
                NoteType, p.Text, p.CustomMark);
        }
        else if (p.ParagraphIndex.HasValue)
        {
            var section = doc.Sections[p.SectionIndex];
            insertedNote = WordNoteHelper.InsertNoteAtParagraph(builder, section, p.ParagraphIndex.Value,
                NoteType, p.Text, p.CustomMark);
        }
        else
        {
            insertedNote = WordNoteHelper.InsertNoteAtDocumentEnd(builder, NoteType, p.Text, p.CustomMark);
        }

        MarkModified(context);

        return BuildSuccessResult(p.Text, insertedNote.ReferenceMark);
    }

    /// <summary>
    ///     Builds the success result for adding a note.
    /// </summary>
    /// <param name="text">The note text content.</param>
    /// <param name="referenceMark">The reference mark assigned to the note.</param>
    /// <returns>A <see cref="SuccessResult" /> containing the operation details.</returns>
    private SuccessResult BuildSuccessResult(string text, string? referenceMark)
    {
        var result = new StringBuilder();
        result.AppendLine($"{NoteTypeName} added successfully");
        result.AppendLine($"Text: {text}");
        if (!string.IsNullOrEmpty(referenceMark))
            result.AppendLine($"Reference mark: {referenceMark}");
        return new SuccessResult { Message = result.ToString().TrimEnd() };
    }

    /// <summary>
    ///     Extracts add note parameters from operation parameters.
    /// </summary>
    /// <param name="parameters">The operation parameters to extract from.</param>
    /// <returns>The extracted <see cref="AddNoteParameters" />.</returns>
    private static AddNoteParameters ExtractAddNoteParameters(OperationParameters parameters)
    {
        return new AddNoteParameters(
            parameters.GetRequired<string>("text"),
            parameters.GetOptional<int?>("paragraphIndex"),
            parameters.GetOptional("sectionIndex", 0),
            parameters.GetOptional<string?>("referenceText"),
            parameters.GetOptional<string?>("customMark")
        );
    }

    /// <summary>
    ///     Record to hold add note parameters.
    /// </summary>
    /// <param name="Text">The note text content.</param>
    /// <param name="ParagraphIndex">Optional paragraph index to insert at.</param>
    /// <param name="SectionIndex">The section index (default: 0).</param>
    /// <param name="ReferenceText">Optional reference text to find insertion point.</param>
    /// <param name="CustomMark">Optional custom reference mark.</param>
    private sealed record AddNoteParameters(
        string Text,
        int? ParagraphIndex,
        int SectionIndex,
        string? ReferenceText,
        string? CustomMark);
}
