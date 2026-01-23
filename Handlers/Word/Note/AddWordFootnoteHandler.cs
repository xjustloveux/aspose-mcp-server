using System.Text;
using Aspose.Words;
using Aspose.Words.Notes;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Helpers.Word;
using AsposeMcpServer.Results.Common;

namespace AsposeMcpServer.Handlers.Word.Note;

/// <summary>
///     Handler for adding footnotes to Word documents.
/// </summary>
[ResultType(typeof(SuccessResult))]
public class AddWordFootnoteHandler : OperationHandlerBase<Document>
{
    /// <inheritdoc />
    public override string Operation => "add_footnote";

    /// <summary>
    ///     Adds a footnote to the document.
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="parameters">
    ///     Required: text
    ///     Optional: paragraphIndex, sectionIndex, referenceText, customMark
    /// </param>
    /// <returns>Success message with footnote details.</returns>
    public override object Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var p = ExtractAddFootnoteParameters(parameters);

        var doc = context.Document;

        if (p.SectionIndex < 0 || p.SectionIndex >= doc.Sections.Count)
            throw new ArgumentException($"sectionIndex must be between 0 and {doc.Sections.Count - 1}");

        var builder = new DocumentBuilder(doc);
        Footnote insertedNote;

        if (!string.IsNullOrEmpty(p.ReferenceText))
        {
            insertedNote = WordNoteHelper.InsertNoteAtReferenceText(doc, builder, p.ReferenceText,
                FootnoteType.Footnote, p.Text, p.CustomMark);
        }
        else if (p.ParagraphIndex.HasValue)
        {
            var section = doc.Sections[p.SectionIndex];
            insertedNote = WordNoteHelper.InsertNoteAtParagraph(builder, section, p.ParagraphIndex.Value,
                FootnoteType.Footnote, p.Text, p.CustomMark);
        }
        else
        {
            insertedNote = WordNoteHelper.InsertNoteAtDocumentEnd(builder, FootnoteType.Footnote, p.Text, p.CustomMark);
        }

        MarkModified(context);

        return BuildSuccessResult(p.Text, insertedNote.ReferenceMark);
    }

    /// <summary>
    ///     Builds the success result for adding a footnote.
    /// </summary>
    /// <param name="text">The footnote text.</param>
    /// <param name="referenceMark">The reference mark.</param>
    /// <returns>The success result.</returns>
    private static SuccessResult BuildSuccessResult(string text, string? referenceMark)
    {
        var result = new StringBuilder();
        result.AppendLine("Footnote added successfully");
        result.AppendLine($"Text: {text}");
        if (!string.IsNullOrEmpty(referenceMark))
            result.AppendLine($"Reference mark: {referenceMark}");
        return new SuccessResult { Message = result.ToString().TrimEnd() };
    }

    /// <summary>
    ///     Extracts add footnote parameters from operation parameters.
    /// </summary>
    /// <param name="parameters">The operation parameters.</param>
    /// <returns>The extracted add footnote parameters.</returns>
    private static AddFootnoteParameters ExtractAddFootnoteParameters(OperationParameters parameters)
    {
        return new AddFootnoteParameters(
            parameters.GetRequired<string>("text"),
            parameters.GetOptional<int?>("paragraphIndex"),
            parameters.GetOptional("sectionIndex", 0),
            parameters.GetOptional<string?>("referenceText"),
            parameters.GetOptional<string?>("customMark")
        );
    }

    /// <summary>
    ///     Record to hold add footnote parameters.
    /// </summary>
    private sealed record AddFootnoteParameters(
        string Text,
        int? ParagraphIndex,
        int SectionIndex,
        string? ReferenceText,
        string? CustomMark);
}
