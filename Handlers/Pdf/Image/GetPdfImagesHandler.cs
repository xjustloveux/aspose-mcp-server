using Aspose.Pdf;
using AsposeMcpServer.Core.Handlers;

namespace AsposeMcpServer.Handlers.Pdf.Image;

/// <summary>
///     Handler for getting image information from PDF documents.
/// </summary>
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
    /// <returns>JSON string containing image information.</returns>
    public override string Execute(OperationContext<Document> context, OperationParameters parameters)
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
    /// <returns>JSON string containing image information from the specified page.</returns>
    private string GetImagesFromSinglePage(Document document, int pageIndex)
    {
        if (pageIndex < 1 || pageIndex > document.Pages.Count)
            throw new ArgumentException($"pageIndex must be between 1 and {document.Pages.Count}");

        var page = document.Pages[pageIndex];
        var images = page.Resources?.Images;

        if (images == null || images.Count == 0)
            return JsonResult(new
            {
                count = 0,
                pageIndex,
                items = Array.Empty<object>(),
                message = $"No images found on page {pageIndex}"
            });

        var imageList = CollectImagesFromPage(images, pageIndex);
        return JsonResult(new
        {
            count = imageList.Count,
            pageIndex,
            items = imageList
        });
    }

    /// <summary>
    ///     Retrieves images from all pages in the document.
    /// </summary>
    /// <param name="document">The PDF document.</param>
    /// <returns>JSON string containing image information from all pages.</returns>
    private string GetImagesFromAllPages(Document document)
    {
        List<object> imageList = [];

        for (var pageNum = 1; pageNum <= document.Pages.Count; pageNum++)
        {
            var page = document.Pages[pageNum];
            var images = page.Resources?.Images;
            if (images is { Count: > 0 })
                imageList.AddRange(CollectImagesFromPage(images, pageNum));
        }

        if (imageList.Count == 0)
            return JsonResult(new
            {
                count = 0,
                items = Array.Empty<object>(),
                message = "No images found in document"
            });

        return JsonResult(new
        {
            count = imageList.Count,
            items = imageList
        });
    }

    /// <summary>
    ///     Collects image information from a page's image collection.
    /// </summary>
    /// <param name="images">The image collection from the page.</param>
    /// <param name="pageNum">The 1-based page number.</param>
    /// <returns>A list of image information objects.</returns>
    private static List<object> CollectImagesFromPage(XImageCollection images, int pageNum)
    {
        List<object> imageList = [];

        for (var i = 1; i <= images.Count; i++)
            try
            {
                var image = images[i];
                var imageInfo = CreateImageInfo(image, i, pageNum);
                imageList.Add(imageInfo);
            }
            catch (Exception ex)
            {
                imageList.Add(new { index = i, pageIndex = pageNum, error = ex.Message });
            }

        return imageList;
    }

    /// <summary>
    ///     Creates an image information dictionary.
    /// </summary>
    /// <param name="image">The XImage object.</param>
    /// <param name="index">The 1-based image index within the page.</param>
    /// <param name="pageNum">The 1-based page number.</param>
    /// <returns>A dictionary containing image information.</returns>
    private static Dictionary<string, object?> CreateImageInfo(XImage image, int index, int pageNum)
    {
        var imageInfo = new Dictionary<string, object?>
        {
            ["index"] = index,
            ["pageIndex"] = pageNum
        };

        try
        {
            if (image is { Width: > 0, Height: > 0 })
            {
                imageInfo["width"] = image.Width;
                imageInfo["height"] = image.Height;
            }
        }
        catch
        {
            imageInfo["width"] = null;
            imageInfo["height"] = null;
        }

        return imageInfo;
    }

    /// <summary>
    ///     Parameters for getting images.
    /// </summary>
    /// <param name="PageIndex">The optional 1-based page index.</param>
    private record GetParameters(int? PageIndex);
}
