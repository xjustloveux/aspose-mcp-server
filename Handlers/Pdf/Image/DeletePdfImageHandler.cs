using Aspose.Pdf;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Results.Common;

namespace AsposeMcpServer.Handlers.Pdf.Image;

/// <summary>
///     Handler for deleting images from PDF documents.
/// </summary>
[ResultType(typeof(SuccessResult))]
public class DeletePdfImageHandler : OperationHandlerBase<Document>
{
    /// <inheritdoc />
    public override string Operation => "delete";

    /// <summary>
    ///     Deletes an image from the specified page of the PDF document.
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="parameters">
    ///     Optional: pageIndex (default: 1), imageIndex (default: 1)
    /// </param>
    /// <returns>Success message with delete details.</returns>
    public override object Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var p = ExtractDeleteParameters(parameters);

        var document = context.Document;

        // Reject non-positive indices instead of clamping to the first page/image: 'get' treats
        // pageIndex 0 as "all pages", so a silent clamp would delete from a page the caller
        // never named.
        if (p.PageIndex < 1 || p.PageIndex > document.Pages.Count)
            throw new ArgumentException($"pageIndex must be between 1 and {document.Pages.Count}");

        var page = document.Pages[p.PageIndex];
        var images = page.Resources?.Images;
        if (images == null)
            throw new ArgumentException("No images found on the page");
        if (p.ImageIndex < 1 || p.ImageIndex > images.Count)
            throw new ArgumentException($"imageIndex must be between 1 and {images.Count}");

        images.Delete(p.ImageIndex);

        MarkModified(context);

        return new SuccessResult { Message = $"Deleted image {p.ImageIndex} from page {p.PageIndex}." };
    }

    /// <summary>
    ///     Extracts delete parameters from the operation parameters.
    /// </summary>
    /// <param name="parameters">The operation parameters.</param>
    /// <returns>The extracted parameters.</returns>
    private static DeleteParameters ExtractDeleteParameters(OperationParameters parameters)
    {
        return new DeleteParameters(
            parameters.GetOptional("pageIndex", 1),
            parameters.GetOptional("imageIndex", 1));
    }

    /// <summary>
    ///     Parameters for deleting an image.
    /// </summary>
    /// <param name="PageIndex">The 1-based page index.</param>
    /// <param name="ImageIndex">The 1-based image index.</param>
    private sealed record DeleteParameters(int PageIndex, int ImageIndex);
}
