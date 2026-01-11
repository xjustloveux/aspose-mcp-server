using Aspose.Words;
using Aspose.Words.Layout;
using AsposeMcpServer.Core.Handlers;
using WordParagraph = Aspose.Words.Paragraph;

namespace AsposeMcpServer.Handlers.Word.Page;

/// <summary>
///     Handler for inserting a blank page at a specified position in Word documents.
/// </summary>
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
    public override string Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var insertAtPageIndex = parameters.GetOptional<int?>("insertAtPageIndex");

        var doc = context.Document;
        var builder = new DocumentBuilder(doc);

        if (insertAtPageIndex is > 0)
        {
            var pageCount = doc.PageCount;
            if (insertAtPageIndex.Value > pageCount)
                throw new ArgumentException(
                    $"insertAtPageIndex must be between 0 and {pageCount} (document has {pageCount} pages)");

            var layoutCollector = new LayoutCollector(doc);
            var paragraphs = doc.GetChildNodes(NodeType.Paragraph, true).Cast<WordParagraph>().ToList();

            WordParagraph? targetParagraph = null;
            foreach (var para in paragraphs)
            {
                var paraPage = layoutCollector.GetStartPageIndex(para);
                if (paraPage == insertAtPageIndex.Value + 1)
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

        return Success($"Blank page inserted at page {insertAtPageIndex ?? doc.PageCount}");
    }
}
