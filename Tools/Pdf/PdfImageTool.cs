using System.ComponentModel;
using System.Drawing.Imaging;
using System.Text.Json;
using Aspose.Pdf;
using AsposeMcpServer.Core.Helpers;
using AsposeMcpServer.Core.Session;
using ModelContextProtocol.Server;

namespace AsposeMcpServer.Tools.Pdf;

/// <summary>
///     Tool for managing images in PDF documents (add, delete, edit, extract, get)
/// </summary>
[McpServerToolType]
public class PdfImageTool
{
    /// <summary>
    ///     The document session manager for managing in-memory document sessions.
    /// </summary>
    private readonly DocumentSessionManager? _sessionManager;

    /// <summary>
    ///     Initializes a new instance of the <see cref="PdfImageTool" /> class.
    /// </summary>
    /// <param name="sessionManager">Optional session manager for in-memory document editing.</param>
    public PdfImageTool(DocumentSessionManager? sessionManager = null)
    {
        _sessionManager = sessionManager;
    }

    [McpServerTool(Name = "pdf_image")]
    [Description(@"Manage images in PDF documents. Supports 5 operations: add, delete, edit, extract, get.

Usage examples:
- Add image: pdf_image(operation='add', path='doc.pdf', pageIndex=1, imagePath='image.png', x=100, y=100)
- Delete image: pdf_image(operation='delete', path='doc.pdf', pageIndex=1, imageIndex=1)
- Move image: pdf_image(operation='edit', path='doc.pdf', pageIndex=1, imageIndex=1, x=200, y=200)
- Replace image: pdf_image(operation='edit', path='doc.pdf', pageIndex=1, imageIndex=1, imagePath='new.png', x=200, y=200)
- Extract image: pdf_image(operation='extract', path='doc.pdf', pageIndex=1, imageIndex=1, outputPath='image.png')
- Get images: pdf_image(operation='get', path='doc.pdf', pageIndex=1)")]
    public string Execute(
        [Description(@"Operation to perform.
- 'add': Add an image (required params: path, pageIndex, imagePath)
- 'delete': Delete an image (required params: path, pageIndex, imageIndex)
- 'edit': Edit image position/size (required params: path, pageIndex, imageIndex)
- 'extract': Extract an image (required params: path, pageIndex, imageIndex, outputPath)
- 'get': Get all images on a page (required params: path, pageIndex)")]
        string operation,
        [Description("PDF file path (required if no sessionId)")]
        string? path = null,
        [Description("Session ID for in-memory editing")]
        string? sessionId = null,
        [Description(
            "Output file path (optional, defaults to overwrite input for add/delete/edit, required for extract)")]
        string? outputPath = null,
        [Description("Page index (1-based, required for add, delete, edit, extract, get)")]
        int pageIndex = 0,
        [Description("Image file path (required for add, optional for edit - omit to move existing image)")]
        string? imagePath = null,
        [Description("Image index (1-based, required for delete, edit, extract)")]
        int imageIndex = 0,
        [Description("X position in PDF coordinates, origin at bottom-left corner (for add, edit, default: 100)")]
        double x = 100,
        [Description("Y position in PDF coordinates, origin at bottom-left corner (for add, edit, default: 600)")]
        double y = 600,
        [Description("Image width (for add, edit, optional - if omitted defaults to 200)")]
        double? width = null,
        [Description("Image height (for add, edit, optional - if omitted defaults to 200)")]
        double? height = null,
        [Description("Output directory for extracted images (for extract)")]
        string? outputDir = null)
    {
        return operation.ToLower() switch
        {
            "add" => AddImage(sessionId, path, outputPath, pageIndex, imagePath, x, y, width, height),
            "delete" => DeleteImage(sessionId, path, outputPath, pageIndex, imageIndex),
            "edit" => EditImage(sessionId, path, outputPath, pageIndex, imageIndex, imagePath, x, y, width, height),
            "extract" => ExtractImages(path, outputPath, outputDir, pageIndex, imageIndex),
            "get" => GetImages(sessionId, path, pageIndex),
            _ => throw new ArgumentException($"Unknown operation: {operation}")
        };
    }

    /// <summary>
    ///     Adds an image to the specified page of the PDF document.
    /// </summary>
    /// <param name="sessionId">The session ID for in-memory editing.</param>
    /// <param name="path">The PDF file path.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="pageIndex">The 1-based page index.</param>
    /// <param name="imagePath">The path to the image file to add.</param>
    /// <param name="x">The X position in PDF coordinates.</param>
    /// <param name="y">The Y position in PDF coordinates.</param>
    /// <param name="width">Optional width for the image.</param>
    /// <param name="height">Optional height for the image.</param>
    /// <returns>A message indicating the result of the operation.</returns>
    /// <exception cref="ArgumentException">Thrown when required parameters are missing or invalid.</exception>
    /// <exception cref="FileNotFoundException">Thrown when the image file is not found.</exception>
    private string AddImage(string? sessionId, string? path, string? outputPath, int pageIndex, string? imagePath,
        double x, double y, double? width, double? height)
    {
        if (string.IsNullOrEmpty(imagePath))
            throw new ArgumentException("imagePath is required for add operation");

        SecurityHelper.ValidateFilePath(imagePath, "imagePath", true);

        if (!File.Exists(imagePath))
            throw new FileNotFoundException($"Image file not found: {imagePath}");

        using var ctx = DocumentContext<Document>.Create(_sessionManager, sessionId, path);
        var document = ctx.Document;

        var actualPageIndex = pageIndex < 1 ? 1 : pageIndex;
        if (actualPageIndex > document.Pages.Count)
            throw new ArgumentException($"pageIndex must be between 1 and {document.Pages.Count}");

        var page = document.Pages[actualPageIndex];
        page.AddImage(imagePath,
            new Rectangle(x, y, width.HasValue ? x + width.Value : x + 200,
                height.HasValue ? y + height.Value : y + 200));

        ctx.Save(outputPath);

        return $"Added image to page {actualPageIndex}. {ctx.GetOutputMessage(outputPath)}";
    }

    /// <summary>
    ///     Deletes an image from the specified page of the PDF document.
    /// </summary>
    /// <param name="sessionId">The session ID for in-memory editing.</param>
    /// <param name="path">The PDF file path.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="pageIndex">The 1-based page index.</param>
    /// <param name="imageIndex">The 1-based image index.</param>
    /// <returns>A message indicating the result of the operation.</returns>
    /// <exception cref="ArgumentException">Thrown when the page or image index is invalid.</exception>
    private string DeleteImage(string? sessionId, string? path, string? outputPath, int pageIndex, int imageIndex)
    {
        using var ctx = DocumentContext<Document>.Create(_sessionManager, sessionId, path);
        var document = ctx.Document;

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

        ctx.Save(outputPath);

        return $"Deleted image {actualImageIndex} from page {actualPageIndex}. {ctx.GetOutputMessage(outputPath)}";
    }

    /// <summary>
    ///     Edits an existing image on the specified page (move or replace).
    /// </summary>
    /// <param name="sessionId">The session ID for in-memory editing.</param>
    /// <param name="path">The PDF file path.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="pageIndex">The 1-based page index.</param>
    /// <param name="imageIndex">The 1-based image index.</param>
    /// <param name="imagePath">Optional new image path to replace the existing image.</param>
    /// <param name="x">The new X position in PDF coordinates.</param>
    /// <param name="y">The new Y position in PDF coordinates.</param>
    /// <param name="width">Optional new width for the image.</param>
    /// <param name="height">Optional new height for the image.</param>
    /// <returns>A message indicating the result of the operation.</returns>
    /// <exception cref="ArgumentException">Thrown when the page or image index is invalid.</exception>
    private string EditImage(string? sessionId, string? path, string? outputPath, int pageIndex, int imageIndex,
        string? imagePath, double? x, double? y, double? width, double? height)
    {
        using var ctx = DocumentContext<Document>.Create(_sessionManager, sessionId, path);
        var document = ctx.Document;

        var actualPageIndex = pageIndex < 1 ? 1 : pageIndex;
        if (actualPageIndex > document.Pages.Count)
            throw new ArgumentException($"pageIndex must be between 1 and {document.Pages.Count}");

        var page = document.Pages[actualPageIndex];
        var images = page.Resources?.Images;
        if (images == null)
            throw new ArgumentException("No images found on the page");
        if (imageIndex < 1 || imageIndex > images.Count)
            throw new ArgumentException($"imageIndex must be between 1 and {images.Count}");

        string? tempImagePath = null;
        try
        {
            if (string.IsNullOrEmpty(imagePath))
            {
                tempImagePath = Path.Combine(Path.GetTempPath(), $"temp_image_{Guid.NewGuid()}.png");
                using var imageStream = new FileStream(tempImagePath, FileMode.Create);
#pragma warning disable CA1416
                images[imageIndex].Save(imageStream, ImageFormat.Png);
#pragma warning restore CA1416
                imagePath = tempImagePath;
            }
            else
            {
                SecurityHelper.ValidateFilePath(imagePath, "imagePath", true);
                if (!File.Exists(imagePath))
                    throw new FileNotFoundException($"Image file not found: {imagePath}");
            }

            images.Delete(imageIndex);
            var newX = x ?? 100;
            var newY = y ?? 600;
            page.AddImage(imagePath,
                new Rectangle(newX, newY, width.HasValue ? newX + width.Value : newX + 200,
                    height.HasValue ? newY + height.Value : newY + 200));

            ctx.Save(outputPath);

            var action = tempImagePath != null ? "Moved" : "Replaced";
            return $"{action} image {imageIndex} on page {pageIndex}. {ctx.GetOutputMessage(outputPath)}";
        }
        finally
        {
            if (tempImagePath != null && File.Exists(tempImagePath))
                File.Delete(tempImagePath);
        }
    }

    /// <summary>
    ///     Extracts images from the specified page of the PDF document.
    /// </summary>
    /// <param name="path">The PDF file path.</param>
    /// <param name="outputPath">The output file path for a single image.</param>
    /// <param name="outputDir">The output directory for multiple images.</param>
    /// <param name="pageIndex">The 1-based page index.</param>
    /// <param name="imageIndex">Optional 1-based image index for extracting a specific image.</param>
    /// <returns>A message indicating the result of the operation.</returns>
    /// <exception cref="ArgumentException">Thrown when required parameters are missing or invalid.</exception>
    private static string ExtractImages(string? path, string? outputPath, string? outputDir, int pageIndex,
        int? imageIndex)
    {
        if (string.IsNullOrEmpty(path))
            throw new ArgumentException("path is required for extract operation");

        SecurityHelper.ValidateFilePath(path, "path", true);

        if (!string.IsNullOrEmpty(outputPath))
            SecurityHelper.ValidateFilePath(outputPath, "outputPath", true);
        if (!string.IsNullOrEmpty(outputDir))
            SecurityHelper.ValidateFilePath(outputDir, "outputDir", true);

        var targetDir = outputDir ?? Path.GetDirectoryName(outputPath) ?? Path.GetDirectoryName(path) ?? ".";
        Directory.CreateDirectory(targetDir);

        using var document = new Document(path);
        var actualPageIndex = pageIndex < 1 ? 1 : pageIndex;
        if (actualPageIndex > document.Pages.Count)
            throw new ArgumentException($"pageIndex must be between 1 and {document.Pages.Count}");

        var page = document.Pages[actualPageIndex];
        var images = page.Resources?.Images;
        if (images == null || images.Count == 0)
            return $"No images found on page {pageIndex}.";

        if (imageIndex is > 0)
        {
            if (imageIndex.Value < 1 || imageIndex.Value > images.Count)
                throw new ArgumentException($"imageIndex must be between 1 and {images.Count}");

            var image = images[imageIndex.Value];
            var fileName = outputPath ?? Path.Combine(targetDir, $"page_{pageIndex}_image_{imageIndex.Value}.png");
            using var imageStream = new FileStream(fileName, FileMode.Create);
#pragma warning disable CA1416
            image.Save(imageStream, ImageFormat.Png);
#pragma warning restore CA1416
            return $"Extracted image {imageIndex.Value} from page {pageIndex} to: {fileName}";
        }

        var count = 0;
        for (var i = 1; i <= images.Count; i++)
        {
            var image = images[i];
            var fileName = Path.Combine(targetDir, $"page_{pageIndex}_image_{i}.png");
            using var imageStream = new FileStream(fileName, FileMode.Create);
#pragma warning disable CA1416
            image.Save(imageStream, ImageFormat.Png);
#pragma warning restore CA1416
            count++;
        }

        return $"Extracted {count} image(s) from page {pageIndex} to: {targetDir}";
    }

    /// <summary>
    ///     Retrieves information about images in the PDF document.
    /// </summary>
    /// <param name="sessionId">The session ID for in-memory editing.</param>
    /// <param name="path">The PDF file path.</param>
    /// <param name="pageIndex">Optional 1-based page index to get images from a specific page.</param>
    /// <returns>A JSON string containing image information.</returns>
    /// <exception cref="ArgumentException">Thrown when the page index is out of range.</exception>
    private string GetImages(string? sessionId, string? path, int? pageIndex)
    {
        using var ctx = DocumentContext<Document>.Create(_sessionManager, sessionId, path);
        var document = ctx.Document;
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
                return JsonSerializer.Serialize(emptyResult, new JsonSerializerOptions { WriteIndented = true });
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
            return JsonSerializer.Serialize(result, new JsonSerializerOptions { WriteIndented = true });
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
                return JsonSerializer.Serialize(emptyResult, new JsonSerializerOptions { WriteIndented = true });
            }

            var result = new
            {
                count = imageList.Count,
                items = imageList
            };
            return JsonSerializer.Serialize(result, new JsonSerializerOptions { WriteIndented = true });
        }
    }
}