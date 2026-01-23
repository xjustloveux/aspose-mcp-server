using Aspose.Words;
using Aspose.Words.Layout;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Results.Common;
using WordParagraph = Aspose.Words.Paragraph;

namespace AsposeMcpServer.Handlers.Word.Page;

/// <summary>
///     Handler for inserting a blank page at a specified position in Word documents.
/// </summary>
[ResultType(typeof(SuccessResult))]
public class InsertBlankPageWordHandler : OperationHandlerBase<Document>
{
    /// <inheritdoc />
    public override string Operation => "insert_blank_page";

    /// <summary>
    ///     Inserts a blank page at the specified position in the document.
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="parameters">
    ///     Optional: insertAtPageIndex (0-based page index to insert at)
    /// </param>
    /// <returns>Success message with insertion details.</returns>
    public override object Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var insertParams = ExtractInsertBlankPageParameters(parameters);

        var doc = context.Document;
        var builder = new DocumentBuilder(doc);

        if (insertParams.InsertAtPageIndex is > 0)
        {
            var pageCount = doc.PageCount;
            if (insertParams.InsertAtPageIndex.Value > pageCount)
                throw new ArgumentException(
                    $"insertAtPageIndex must be between 0 and {pageCount} (document has {pageCount} pages)");

            var layoutCollector = new LayoutCollector(doc);
            var paragraphs = doc.GetChildNodes(NodeType.Paragraph, true).Cast<WordParagraph>().ToList();

            WordParagraph? targetParagraph = null;
            foreach (var para in paragraphs)
            {
                var paraPage = layoutCollector.GetStartPageIndex(para);
                if (paraPage == insertParams.InsertAtPageIndex.Value + 1)
                {
                    targetParagraph = para;
                    break;
                }
            }

            if (targetParagraph != null)
            {
                builder.MoveTo(targetParagraph);
                builder.InsertBreak(BreakType.PageBreak);
            }
            else
            {
                builder.MoveToDocumentEnd();
                builder.InsertBreak(BreakType.PageBreak);
            }
        }
        else
        {
            builder.MoveToDocumentEnd();
            builder.InsertBreak(BreakType.PageBreak);
        }

        MarkModified(context);

        return new SuccessResult
            { Message = $"Blank page inserted at page {insertParams.InsertAtPageIndex ?? doc.PageCount}" };
    }

    /// <summary>
    ///     Extracts insert blank page parameters from operation parameters.
    /// </summary>
    /// <param name="parameters">The operation parameters.</param>
    /// <returns>The extracted insert blank page parameters.</returns>
    private static InsertBlankPageParameters ExtractInsertBlankPageParameters(OperationParameters parameters)
    {
        return new InsertBlankPageParameters(
            parameters.GetOptional<int?>("insertAtPageIndex")
        );
    }

    /// <summary>
    ///     Record to hold insert blank page parameters.
    /// </summary>
    /// <param name="InsertAtPageIndex">The 0-based page index to insert at.</param>
    private sealed record InsertBlankPageParameters(int? InsertAtPageIndex);
}
