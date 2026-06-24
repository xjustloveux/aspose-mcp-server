using Aspose.Words;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Helpers.Word;
using AsposeMcpServer.Results.Common;

namespace AsposeMcpServer.Handlers.Word.SectionBreak;

/// <summary>
///     Handler for inserting section breaks in Word documents.
/// </summary>
[ResultType(typeof(SuccessResult))]
public class InsertWordSectionHandler : OperationHandlerBase<Document>
{
    /// <inheritdoc />
    public override string Operation => "insert";

    /// <summary>
    ///     Inserts a section break into the document at specified position.
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="parameters">
    ///     Required: sectionBreakType
    ///     Optional: insertAtParagraphIndex, sectionIndex
    /// </param>
    /// <returns>Success message with section insertion details.</returns>
    public override object Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var p = ExtractInsertWordSectionParameters(parameters);

        if (string.IsNullOrEmpty(p.SectionBreakType))
            throw new ArgumentException("sectionBreakType is required for insert operation");

        var doc = context.Document;
        var builder = new DocumentBuilder(doc);

        var breakType = WordSectionHelper.GetSectionStart(p.SectionBreakType);

        if (p.InsertAtParagraphIndex.HasValue && p.InsertAtParagraphIndex.Value != -1)
        {
            var para = ParagraphResolver
                .Resolve(doc, ParagraphAddress.From(parameters, p.InsertAtParagraphIndex.Value)).Paragraph;
            builder.MoveTo(para);
        }
        else
        {
            builder.MoveToDocumentEnd();
        }

        builder.InsertBreak(BreakType.SectionBreakContinuous);
        builder.CurrentSection.PageSetup.SectionStart = breakType;

        MarkModified(context);

        return new SuccessResult { Message = $"Section break inserted ({p.SectionBreakType})" };
    }

    private static InsertWordSectionParameters ExtractInsertWordSectionParameters(OperationParameters parameters)
    {
        return new InsertWordSectionParameters(
            parameters.GetRequired<string>("sectionBreakType"),
            parameters.GetOptional<int?>("insertAtParagraphIndex"));
    }

    private sealed record InsertWordSectionParameters(
        string SectionBreakType,
        int? InsertAtParagraphIndex);
}
