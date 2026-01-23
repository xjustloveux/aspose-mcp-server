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

        var actualPageIndex = p.PageIndex < 1 ? 1 : p.PageIndex;
        var actualImageIndex = p.ImageIndex < 1 ? 1 : p.ImageIndex;
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

        return new SuccessResult { Message = $"Deleted image {actualImageIndex} from page {actualPageIndex}." };
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
