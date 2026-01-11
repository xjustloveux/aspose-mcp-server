using Aspose.Pdf;
using AsposeMcpServer.Core.Handlers;

namespace AsposeMcpServer.Handlers.Pdf.Image;

/// <summary>
///     Handler for deleting images from PDF documents.
/// </summary>
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
    public override string Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var pageIndex = parameters.GetOptional("pageIndex", 1);
        var imageIndex = parameters.GetOptional("imageIndex", 1);

        var document = context.Document;

        var actualPageIndex = pageIndex < 1 ? 1 : pageIndex;
        var actualImageIndex = imageIndex < 1 ? 1 : imageIndex;
        if (actualPageIndex > document.Pages.Count)
            throw new ArgumentException($"pageIndex must be between 1 and {document.Pages.Count}");

        var page = document.Pages[actualPageIndex];
        var images = page.Resources?.Images;
        if (images == null)
            throw new ArgumentException("No images found on the page");
        if (actualImageIndex > images.Count)
            throw new ArgumentException($"imageIndex must be between 1 and {images.Count}");

        images.Delete(actualImageIndex);

        MarkModified(context);

        return Success($"Deleted image {actualImageIndex} from page {actualPageIndex}.");
    }
}
