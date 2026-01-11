using Aspose.Pdf;
using Aspose.Pdf.Annotations;
using AsposeMcpServer.Core.Handlers;

namespace AsposeMcpServer.Handlers.Pdf.Link;

/// <summary>
///     Handler for deleting links from PDF documents.
/// </summary>
public class DeletePdfLinkHandler : OperationHandlerBase<Document>
{
    /// <inheritdoc />
    public override string Operation => "delete";

    /// <summary>
    ///     Deletes a link from a specific page in the PDF document.
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="parameters">
    ///     Required: pageIndex, linkIndex.
    /// </param>
    /// <returns>Success message with deletion details.</returns>
    public override string Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var pageIndex = parameters.GetRequired<int>("pageIndex");
        var linkIndex = parameters.GetRequired<int>("linkIndex");

        var document = context.Document;

        if (pageIndex < 1 || pageIndex > document.Pages.Count)
            throw new ArgumentException($"pageIndex must be between 1 and {document.Pages.Count}");

        var page = document.Pages[pageIndex];

        var linkAnnotations = page.Annotations
            .Where(a => a is LinkAnnotation)
            .ToList();

        if (linkIndex < 0 || linkIndex >= linkAnnotations.Count)
            throw new ArgumentException(
                $"linkIndex must be between 0 and {linkAnnotations.Count - 1}");

        page.Annotations.Delete(linkAnnotations[linkIndex]);

        MarkModified(context);

        return Success($"Link {linkIndex} deleted from page {pageIndex}.");
    }
}
