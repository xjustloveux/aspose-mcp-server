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
        List<object> imageList = [];

        if (pageIndex is > 0)
        {
            if (pageIndex.Value < 1 || pageIndex.Value > document.Pages.Count)
                throw new ArgumentException($"pageIndex must be between 1 and {document.Pages.Count}");
            var page = document.Pages[pageIndex.Value];
            var images = page.Resources?.Images;

            if (images == null || images.Count == 0)
            {
                var emptyResult = new
                {
                    count = 0,
                    pageIndex = pageIndex.Value,
                    items = Array.Empty<object>(),
                    message = $"No images found on page {pageIndex.Value}"
                };
                return JsonResult(emptyResult);
            }

            for (var i = 1; i <= images.Count; i++)
                try
                {
                    var image = images[i];
                    var imageInfo = new Dictionary<string, object?>
                    {
                        ["index"] = i,
                        ["pageIndex"] = pageIndex.Value
                    };
                    try
                    {
                        if (image.Width > 0 && image.Height > 0)
                        {
                            imageInfo["width"] = image.Width;
                            imageInfo["height"] = image.Height;
                        }
                    }
                    catch (Exception ex)
                    {
                        imageInfo["width"] = null;
                        imageInfo["height"] = null;
                        Console.Error.WriteLine($"[WARN] Failed to read image size: {ex.Message}");
                    }

                    imageList.Add(imageInfo);
                }
                catch (Exception ex)
                {
                    imageList.Add(new { index = i, pageIndex = pageIndex.Value, error = ex.Message });
                }

            var result = new
            {
                count = imageList.Count,
                pageIndex = pageIndex.Value,
                items = imageList
            };
            return JsonResult(result);
        }
        else
        {
            for (var pageNum = 1; pageNum <= document.Pages.Count; pageNum++)
            {
                var page = document.Pages[pageNum];
                var images = page.Resources?.Images;
                if (images is { Count: > 0 })
                    for (var i = 1; i <= images.Count; i++)
                        try
                        {
                            var image = images[i];
                            var imageInfo = new Dictionary<string, object?>
                            {
                                ["index"] = i,
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

                            imageList.Add(imageInfo);
                        }
                        catch (Exception ex)
                        {
                            imageList.Add(new { index = i, pageIndex = pageNum, error = ex.Message });
                        }
            }

            if (imageList.Count == 0)
            {
                var emptyResult = new
                {
                    count = 0,
                    items = Array.Empty<object>(),
                    message = "No images found in document"
                };
                return JsonResult(emptyResult);
            }

            var result = new
            {
                count = imageList.Count,
                items = imageList
            };
            return JsonResult(result);
        }
    }
}
