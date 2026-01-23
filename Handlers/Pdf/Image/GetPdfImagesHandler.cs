using Aspose.Pdf;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Results.Pdf.Image;

namespace AsposeMcpServer.Handlers.Pdf.Image;

/// <summary>
///     Handler for getting image information from PDF documents.
/// </summary>
[ResultType(typeof(GetImagesPdfResult))]
public class GetPdfImagesHandler : OperationHandlerBase<Document>
{
    /// <inheritdoc />
    public override string Operation => "get";

    /// <summary>
    ///     Retrieves information about images in the PDF document.
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="parameters">
    ///     Optional: pageIndex
    /// </param>
    /// <returns>Result containing image information.</returns>
    public override object Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var p = ExtractGetParameters(parameters);
        var document = context.Document;

        if (p.PageIndex is > 0)
            return GetImagesFromSinglePage(document, p.PageIndex.Value);

        return GetImagesFromAllPages(document);
    }

    /// <summary>
    ///     Extracts get parameters from the operation parameters.
    /// </summary>
    /// <param name="parameters">The operation parameters.</param>
    /// <returns>The extracted parameters.</returns>
    private static GetParameters ExtractGetParameters(OperationParameters parameters)
    {
        return new GetParameters(
            parameters.GetOptional<int?>("pageIndex"));
    }

    /// <summary>
    ///     Retrieves images from a single page.
    /// </summary>
    /// <param name="document">The PDF document.</param>
    /// <param name="pageIndex">The 1-based page index.</param>
    /// <returns>Result containing image information from the specified page.</returns>
    private static GetImagesPdfResult GetImagesFromSinglePage(Document document, int pageIndex)
    {
        if (pageIndex < 1 || pageIndex > document.Pages.Count)
            throw new ArgumentException($"pageIndex must be between 1 and {document.Pages.Count}");

        var page = document.Pages[pageIndex];
        var images = page.Resources?.Images;

        if (images == null || images.Count == 0)
            return new GetImagesPdfResult
            {
                Count = 0,
                PageIndex = pageIndex,
                Items = [],
                Message = $"No images found on page {pageIndex}"
            };

        var imageList = CollectImagesFromPage(images, pageIndex);
        return new GetImagesPdfResult
        {
            Count = imageList.Count,
            PageIndex = pageIndex,
            Items = imageList
        };
    }

    /// <summary>
    ///     Retrieves images from all pages in the document.
    /// </summary>
    /// <param name="document">The PDF document.</param>
    /// <returns>Result containing image information from all pages.</returns>
    private static GetImagesPdfResult GetImagesFromAllPages(Document document)
    {
        List<PdfImageInfo> imageList = [];

        for (var pageNum = 1; pageNum <= document.Pages.Count; pageNum++)
        {
            var page = document.Pages[pageNum];
            var images = page.Resources?.Images;
            if (images is { Count: > 0 })
                imageList.AddRange(CollectImagesFromPage(images, pageNum));
        }

        if (imageList.Count == 0)
            return new GetImagesPdfResult
            {
                Count = 0,
                Items = [],
                Message = "No images found in document"
            };

        return new GetImagesPdfResult
        {
            Count = imageList.Count,
            Items = imageList
        };
    }

    /// <summary>
    ///     Collects image information from a page's image collection.
    /// </summary>
    /// <param name="images">The image collection from the page.</param>
    /// <param name="pageNum">The 1-based page number.</param>
    /// <returns>A list of image information objects.</returns>
    private static List<PdfImageInfo> CollectImagesFromPage(XImageCollection images, int pageNum)
    {
        List<PdfImageInfo> imageList = [];

        for (var i = 1; i <= images.Count; i++)
            try
            {
                var image = images[i];
                var imageInfo = CreateImageInfo(image, i, pageNum);
                imageList.Add(imageInfo);
            }
            catch (Exception ex)
            {
                imageList.Add(new PdfImageInfo { Index = i, PageIndex = pageNum, Error = ex.Message });
            }

        return imageList;
    }

    /// <summary>
    ///     Creates an image information record.
    /// </summary>
    /// <param name="image">The XImage object.</param>
    /// <param name="index">The 1-based image index within the page.</param>
    /// <param name="pageNum">The 1-based page number.</param>
    /// <returns>A PdfImageInfo containing image information.</returns>
    private static PdfImageInfo CreateImageInfo(XImage image, int index, int pageNum)
    {
        int? width = null;
        int? height = null;

        try
        {
            if (image is { Width: > 0, Height: > 0 })
            {
                width = image.Width;
                height = image.Height;
            }
        }
        catch
        {
            // Ignore dimension retrieval errors
        }

        return new PdfImageInfo
        {
            Index = index,
            PageIndex = pageNum,
            Width = width,
            Height = height
        };
    }

    /// <summary>
    ///     Parameters for getting images.
    /// </summary>
    /// <param name="PageIndex">The optional 1-based page index.</param>
    private sealed record GetParameters(int? PageIndex);
}
