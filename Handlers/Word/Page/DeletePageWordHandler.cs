using Aspose.Words;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Results.Common;

namespace AsposeMcpServer.Handlers.Word.Page;

/// <summary>
///     Handler for deleting a specific page from Word documents.
/// </summary>
[ResultType(typeof(SuccessResult))]
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
    public override object Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var deleteParams = ExtractDeletePageParameters(parameters);

        if (!deleteParams.PageIndex.HasValue)
            throw new ArgumentException("pageIndex parameter is required for delete_page operation");

        var doc = context.Document;
        var pageCount = doc.PageCount;

        if (deleteParams.PageIndex.Value < 0 || deleteParams.PageIndex.Value >= pageCount)
            throw new ArgumentException(
                $"pageIndex must be between 0 and {pageCount - 1} (document has {pageCount} pages)");

        var resultDoc = new Document();
        resultDoc.RemoveAllChildren();

        if (deleteParams.PageIndex.Value > 0)
        {
            var beforePages = doc.ExtractPages(0, deleteParams.PageIndex.Value);
            foreach (var section in beforePages.Sections.Cast<Section>())
                resultDoc.AppendChild(resultDoc.ImportNode(section, true));
        }

        if (deleteParams.PageIndex.Value < pageCount - 1)
        {
            var afterPages = doc.ExtractPages(deleteParams.PageIndex.Value + 1,
                pageCount - deleteParams.PageIndex.Value - 1);
            foreach (var section in afterPages.Sections.Cast<Section>())
                resultDoc.AppendChild(resultDoc.ImportNode(section, true));
        }

        context.ResultDocument = resultDoc;
        MarkModified(context);

        return new SuccessResult
        {
            Message =
                $"Page {deleteParams.PageIndex.Value} deleted successfully (document now has {resultDoc.PageCount} pages)"
        };
    }

    /// <summary>
    ///     Extracts delete page parameters from operation parameters.
    /// </summary>
    /// <param name="parameters">The operation parameters.</param>
    /// <returns>The extracted delete page parameters.</returns>
    private static DeletePageParameters ExtractDeletePageParameters(OperationParameters parameters)
    {
        return new DeletePageParameters(
            parameters.GetOptional<int?>("pageIndex")
        );
    }

    /// <summary>
    ///     Record to hold delete page parameters.
    /// </summary>
    /// <param name="PageIndex">The 0-based page index to delete.</param>
    private sealed record DeletePageParameters(int? PageIndex);
}
