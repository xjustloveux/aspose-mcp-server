using Aspose.Words;
using AsposeMcpServer.Core.Handlers;

namespace AsposeMcpServer.Handlers.Word.Page;

/// <summary>
///     Handler for deleting a specific page from Word documents.
/// </summary>
public class DeletePageWordHandler : OperationHandlerBase<Document>
{
    /// <inheritdoc />
    public override string Operation => "delete_page";

    /// <summary>
    ///     Deletes a specific page from the document by extracting and recombining pages.
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="parameters">
    ///     Required: pageIndex (0-based page index to delete)
    /// </param>
    /// <returns>Success message with page deletion details.</returns>
    public override string Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var pageIndex = parameters.GetOptional<int?>("pageIndex");

        if (!pageIndex.HasValue)
            throw new ArgumentException("pageIndex parameter is required for delete_page operation");

        var doc = context.Document;
        var pageCount = doc.PageCount;

        if (pageIndex.Value < 0 || pageIndex.Value >= pageCount)
            throw new ArgumentException(
                $"pageIndex must be between 0 and {pageCount - 1} (document has {pageCount} pages)");

        var resultDoc = new Document();
        resultDoc.RemoveAllChildren();

        if (pageIndex.Value > 0)
        {
            var beforePages = doc.ExtractPages(0, pageIndex.Value);
            foreach (var section in beforePages.Sections.Cast<Section>())
                resultDoc.AppendChild(resultDoc.ImportNode(section, true));
        }

        if (pageIndex.Value < pageCount - 1)
        {
            var afterPages = doc.ExtractPages(pageIndex.Value + 1, pageCount - pageIndex.Value - 1);
            foreach (var section in afterPages.Sections.Cast<Section>())
                resultDoc.AppendChild(resultDoc.ImportNode(section, true));
        }

        // Store the result document for special handling
        context.ResultDocument = resultDoc;
        MarkModified(context);

        return Success($"Page {pageIndex.Value} deleted successfully (document now has {resultDoc.PageCount} pages)");
    }
}
