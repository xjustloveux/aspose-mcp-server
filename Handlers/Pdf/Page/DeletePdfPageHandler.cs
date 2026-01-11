using Aspose.Pdf;
using AsposeMcpServer.Core.Handlers;

namespace AsposeMcpServer.Handlers.Pdf.Page;

/// <summary>
///     Handler for deleting pages from PDF documents.
/// </summary>
public class DeletePdfPageHandler : OperationHandlerBase<Document>
{
    /// <inheritdoc />
    public override string Operation => "delete";

    /// <summary>
    ///     Deletes a page from the PDF document.
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="parameters">
    ///     Required: pageIndex (1-based page index to delete)
    /// </param>
    /// <returns>Success message.</returns>
    /// <exception cref="ArgumentException">Thrown when pageIndex is out of range.</exception>
    public override string Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var pageIndex = parameters.GetRequired<int>("pageIndex");

        var doc = context.Document;

        if (pageIndex < 1 || pageIndex > doc.Pages.Count)
            throw new ArgumentException($"pageIndex must be between 1 and {doc.Pages.Count}");

        doc.Pages.Delete(pageIndex);

        MarkModified(context);

        return Success($"Deleted page {pageIndex} (remaining: {doc.Pages.Count}).");
    }
}
