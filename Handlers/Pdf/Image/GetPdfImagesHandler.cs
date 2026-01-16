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
        var pageIndex = parameters.GetOptional<int?>("pageIndex");
        var document = context.Document;

        if (pageIndex is > 0)
            return GetImagesFromSinglePage(document, pageIndex.Value);

        return GetImagesFromAllPages(document);
    }

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

    private static Dictionary<string, object?> CreateImageInfo(XImage image, int index, int pageNum)
    {
        var imageInfo = new Dictionary<string, object?>
        {
            ["index"] = index,
            ["pageIndex"] = pageNum
        };

        try
        {
            if (image.Width > 0 && image.Height > 0)
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
}
