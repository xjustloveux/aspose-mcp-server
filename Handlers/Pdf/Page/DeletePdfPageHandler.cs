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
        var p = ExtractDeleteParameters(parameters);

        var doc = context.Document;

        if (p.PageIndex < 1 || p.PageIndex > doc.Pages.Count)
            throw new ArgumentException($"pageIndex must be between 1 and {doc.Pages.Count}");

        doc.Pages.Delete(p.PageIndex);

        MarkModified(context);

        return Success($"Deleted page {p.PageIndex} (remaining: {doc.Pages.Count}).");
    }

    /// <summary>
    ///     Extracts delete parameters from the operation parameters.
    /// </summary>
    /// <param name="parameters">The operation parameters.</param>
    /// <returns>The extracted parameters.</returns>
    private static DeleteParameters ExtractDeleteParameters(OperationParameters parameters)
    {
        return new DeleteParameters(
            parameters.GetRequired<int>("pageIndex"));
    }

    /// <summary>
    ///     Parameters for deleting a page.
    /// </summary>
    /// <param name="PageIndex">The 1-based page index to delete.</param>
    private sealed record DeleteParameters(int PageIndex);
}
