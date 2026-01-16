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
        var p = ExtractDeleteParameters(parameters);

        var document = context.Document;

        if (p.PageIndex < 1 || p.PageIndex > document.Pages.Count)
            throw new ArgumentException($"pageIndex must be between 1 and {document.Pages.Count}");

        var page = document.Pages[p.PageIndex];

        var linkAnnotations = page.Annotations
            .Where(a => a is LinkAnnotation)
            .ToList();

        if (p.LinkIndex < 0 || p.LinkIndex >= linkAnnotations.Count)
            throw new ArgumentException(
                $"linkIndex must be between 0 and {linkAnnotations.Count - 1}");

        page.Annotations.Delete(linkAnnotations[p.LinkIndex]);

        MarkModified(context);

        return Success($"Link {p.LinkIndex} deleted from page {p.PageIndex}.");
    }

    /// <summary>
    ///     Extracts delete parameters from the operation parameters.
    /// </summary>
    /// <param name="parameters">The operation parameters.</param>
    /// <returns>The extracted parameters.</returns>
    private static DeleteParameters ExtractDeleteParameters(OperationParameters parameters)
    {
        return new DeleteParameters(
            parameters.GetRequired<int>("pageIndex"),
            parameters.GetRequired<int>("linkIndex"));
    }

    /// <summary>
    ///     Parameters for deleting a link.
    /// </summary>
    /// <param name="PageIndex">The 1-based page index.</param>
    /// <param name="LinkIndex">The 0-based link index.</param>
    private record DeleteParameters(int PageIndex, int LinkIndex);
}
