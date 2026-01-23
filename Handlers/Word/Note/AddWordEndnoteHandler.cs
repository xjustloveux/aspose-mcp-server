using System.Text;
using Aspose.Words;
using Aspose.Words.Notes;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Helpers.Word;
using AsposeMcpServer.Results.Common;

namespace AsposeMcpServer.Handlers.Word.Note;

/// <summary>
///     Handler for adding endnotes to Word documents.
/// </summary>
[ResultType(typeof(SuccessResult))]
public class AddWordEndnoteHandler : OperationHandlerBase<Document>
{
    /// <inheritdoc />
    public override string Operation => "add_endnote";

    /// <summary>
    ///     Adds an endnote to the document.
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="parameters">
    ///     Required: text
    ///     Optional: paragraphIndex, sectionIndex, referenceText, customMark
    /// </param>
    /// <returns>Success message with endnote details.</returns>
    public override object Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var p = ExtractAddEndnoteParameters(parameters);

        var doc = context.Document;

        if (p.SectionIndex < 0 || p.SectionIndex >= doc.Sections.Count)
            throw new ArgumentException($"sectionIndex must be between 0 and {doc.Sections.Count - 1}");

        var builder = new DocumentBuilder(doc);
        Footnote insertedNote;

        if (!string.IsNullOrEmpty(p.ReferenceText))
        {
            insertedNote = WordNoteHelper.InsertNoteAtReferenceText(doc, builder, p.ReferenceText,
                FootnoteType.Endnote, p.Text, p.CustomMark);
        }
        else if (p.ParagraphIndex.HasValue)
        {
            var section = doc.Sections[p.SectionIndex];
            insertedNote = WordNoteHelper.InsertNoteAtParagraph(builder, section, p.ParagraphIndex.Value,
                FootnoteType.Endnote, p.Text, p.CustomMark);
        }
        else
        {
            insertedNote = WordNoteHelper.InsertNoteAtDocumentEnd(builder, FootnoteType.Endnote, p.Text, p.CustomMark);
        }

        MarkModified(context);

        return BuildSuccessResult(p.Text, insertedNote.ReferenceMark);
    }

    /// <summary>
    ///     Builds the success result for adding an endnote.
    /// </summary>
    /// <param name="text">The endnote text.</param>
    /// <param name="referenceMark">The reference mark.</param>
    /// <returns>The success result.</returns>
    private static SuccessResult BuildSuccessResult(string text, string? referenceMark)
    {
        var result = new StringBuilder();
        result.AppendLine("Endnote added successfully");
        result.AppendLine($"Text: {text}");
        if (!string.IsNullOrEmpty(referenceMark))
            result.AppendLine($"Reference mark: {referenceMark}");
        return new SuccessResult { Message = result.ToString().TrimEnd() };
    }

    /// <summary>
    ///     Extracts add endnote parameters from operation parameters.
    /// </summary>
    /// <param name="parameters">The operation parameters.</param>
    /// <returns>The extracted add endnote parameters.</returns>
    private static AddEndnoteParameters ExtractAddEndnoteParameters(OperationParameters parameters)
    {
        return new AddEndnoteParameters(
            parameters.GetRequired<string>("text"),
            parameters.GetOptional<int?>("paragraphIndex"),
            parameters.GetOptional("sectionIndex", 0),
            parameters.GetOptional<string?>("referenceText"),
            parameters.GetOptional<string?>("customMark")
        );
    }

    /// <summary>
    ///     Record to hold add endnote parameters.
    /// </summary>
    private sealed record AddEndnoteParameters(
        string Text,
        int? ParagraphIndex,
        int SectionIndex,
        string? ReferenceText,
        string? CustomMark);
}
