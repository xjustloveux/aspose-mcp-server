using System.Drawing.Imaging;
using System.Text;
using System.Text.Json.Nodes;
using Aspose.Pdf;
using AsposeMcpServer.Core;

namespace AsposeMcpServer.Tools.Pdf;

/// <summary>
///     Tool for managing images in PDF documents (add, delete, edit, extract)
/// </summary>
public class PdfImageTool : IAsposeTool
{
    public string Description =>
        @"Manage images in PDF documents. Supports 5 operations: add, delete, edit, extract, get.

Usage examples:
- Add image: pdf_image(operation='add', path='doc.pdf', pageIndex=1, imagePath='image.png', x=100, y=100)
- Delete image: pdf_image(operation='delete', path='doc.pdf', pageIndex=1, imageIndex=1)
- Edit image: pdf_image(operation='edit', path='doc.pdf', pageIndex=1, imageIndex=1, x=200, y=200)
- Extract image: pdf_image(operation='extract', path='doc.pdf', pageIndex=1, imageIndex=1, outputPath='image.png')
- Get images: pdf_image(operation='get', path='doc.pdf', pageIndex=1)";

    public object InputSchema => new
    {
        type = "object",
        properties = new
        {
            operation = new
            {
                type = "string",
                description = @"Operation to perform.
- 'add': Add an image (required params: path, pageIndex, imagePath)
- 'delete': Delete an image (required params: path, pageIndex, imageIndex)
- 'edit': Edit image position/size (required params: path, pageIndex, imageIndex)
- 'extract': Extract an image (required params: path, pageIndex, imageIndex, outputPath)
- 'get': Get all images on a page (required params: path, pageIndex)",
                @enum = new[] { "add", "delete", "edit", "extract", "get" }
            },
            path = new
            {
                type = "string",
                description = "PDF file path (required for all operations)"
            },
            outputPath = new
            {
                type = "string",
                description =
                    "Output file path (optional, defaults to overwrite input for add/delete/edit, required for extract)"
            },
            pageIndex = new
            {
                type = "number",
                description = "Page index (1-based, required for add, delete, edit, extract, get)"
            },
            imagePath = new
            {
                type = "string",
                description = "Image file path (required for add, edit)"
            },
            imageIndex = new
            {
                type = "number",
                description = "Image index (1-based, required for delete, edit, extract)"
            },
            x = new
            {
                type = "number",
                description = "X position (for add, edit, default: 100)"
            },
            y = new
            {
                type = "number",
                description = "Y position (for add, edit, default: 600)"
            },
            width = new
            {
                type = "number",
                description = "Image width (for add, edit, optional)"
            },
            height = new
            {
                type = "number",
                description = "Image height (for add, edit, optional)"
            },
            outputDir = new
            {
                type = "string",
                description = "Output directory for extracted images (for extract)"
            }
        },
        required = new[] { "operation", "path" }
    };

    /// <summary>
    ///     Executes the tool operation with the provided JSON arguments
    /// </summary>
    /// <param name="arguments">JSON arguments object containing operation parameters</param>
    /// <returns>Result message as a string</returns>
    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var operation = ArgumentHelper.GetString(arguments, "operation");

        return operation.ToLower() switch
        {
            "add" => await AddImage(arguments),
            "delete" => await DeleteImage(arguments),
            "edit" => await EditImage(arguments),
            "extract" => await ExtractImages(arguments),
            "get" => await GetImages(arguments),
            _ => throw new ArgumentException($"Unknown operation: {operation}")
        };
    }

    /// <summary>
    ///     Adds an image to a PDF page
    /// </summary>
    /// <param name="arguments">JSON arguments containing path, pageIndex, imagePath, x, y, optional width, height, outputPath</param>
    /// <returns>Success message</returns>
    private async Task<string> AddImage(JsonObject? arguments)
    {
        var path = ArgumentHelper.GetAndValidatePath(arguments);
        var outputPath = ArgumentHelper.GetAndValidateOutputPath(arguments, path);
        var pageIndex = ArgumentHelper.GetInt(arguments, "pageIndex");
        var imagePath = ArgumentHelper.GetString(arguments, "imagePath");
        var x = ArgumentHelper.GetDouble(arguments, "x", "x", false, 100);
        var y = ArgumentHelper.GetDouble(arguments, "y", "y", false, 600);
        var width = ArgumentHelper.GetDoubleNullable(arguments, "width");
        var height = ArgumentHelper.GetDoubleNullable(arguments, "height");

        SecurityHelper.ValidateFilePath(path);
        SecurityHelper.ValidateFilePath(outputPath, "outputPath");
        SecurityHelper.ValidateFilePath(imagePath, "imagePath");

        if (!File.Exists(imagePath))
            throw new FileNotFoundException($"Image file not found: {imagePath}");

        using var document = new Document(path);
        if (pageIndex < 1 || pageIndex > document.Pages.Count)
            throw new ArgumentException($"pageIndex must be between 1 and {document.Pages.Count}");

        var page = document.Pages[pageIndex];
        page.AddImage(imagePath,
            new Rectangle(x, y, width.HasValue ? x + width.Value : x + 200,
                height.HasValue ? y + height.Value : y + 200));
        document.Save(outputPath);
        return await Task.FromResult($"Successfully added image to page {pageIndex}. Output: {outputPath}");
    }

    /// <summary>
    ///     Deletes an image from a PDF page
    /// </summary>
    /// <param name="arguments">JSON arguments containing path, pageIndex, imageIndex, optional outputPath</param>
    /// <returns>Success message</returns>
    private async Task<string> DeleteImage(JsonObject? arguments)
    {
        var path = ArgumentHelper.GetAndValidatePath(arguments);
        var outputPath = ArgumentHelper.GetAndValidateOutputPath(arguments, path);
        var pageIndex = ArgumentHelper.GetInt(arguments, "pageIndex");
        var imageIndex = ArgumentHelper.GetInt(arguments, "imageIndex");

        SecurityHelper.ValidateFilePath(path);
        SecurityHelper.ValidateFilePath(outputPath, "outputPath");

        using var document = new Document(path);
        if (pageIndex < 1 || pageIndex > document.Pages.Count)
            throw new ArgumentException($"pageIndex must be between 1 and {document.Pages.Count}");

        var page = document.Pages[pageIndex];
        var images = page.Resources?.Images;
        if (images == null)
            throw new ArgumentException("No images found on the page");
        if (imageIndex < 1 || imageIndex > images.Count)
            throw new ArgumentException($"imageIndex must be between 1 and {images.Count}");

        images.Delete(imageIndex);
        document.Save(outputPath);
        return await Task.FromResult(
            $"Successfully deleted image {imageIndex} from page {pageIndex}. Output: {outputPath}");
    }

    /// <summary>
    ///     Edits image properties in a PDF
    /// </summary>
    /// <param name="arguments">JSON arguments containing path, pageIndex, imageIndex, optional x, y, width, height, outputPath</param>
    /// <returns>Success message</returns>
    private async Task<string> EditImage(JsonObject? arguments)
    {
        var path = ArgumentHelper.GetAndValidatePath(arguments);
        var outputPath = ArgumentHelper.GetAndValidateOutputPath(arguments, path);
        var pageIndex = ArgumentHelper.GetInt(arguments, "pageIndex");
        var imageIndex = ArgumentHelper.GetInt(arguments, "imageIndex");
        var imagePath = ArgumentHelper.GetString(arguments, "imagePath");
        var x = ArgumentHelper.GetDoubleNullable(arguments, "x");
        var y = ArgumentHelper.GetDoubleNullable(arguments, "y");
        var width = ArgumentHelper.GetDoubleNullable(arguments, "width");
        var height = ArgumentHelper.GetDoubleNullable(arguments, "height");

        SecurityHelper.ValidateFilePath(path);
        SecurityHelper.ValidateFilePath(outputPath, "outputPath");
        SecurityHelper.ValidateFilePath(imagePath, "imagePath");

        if (!File.Exists(imagePath))
            throw new FileNotFoundException($"Image file not found: {imagePath}");

        using var document = new Document(path);
        if (pageIndex < 1 || pageIndex > document.Pages.Count)
            throw new ArgumentException($"pageIndex must be between 1 and {document.Pages.Count}");

        var page = document.Pages[pageIndex];
        var images = page.Resources?.Images;
        if (images == null)
            throw new ArgumentException("No images found on the page");
        if (imageIndex < 1 || imageIndex > images.Count)
            throw new ArgumentException($"imageIndex must be between 1 and {images.Count}");

        images.Delete(imageIndex);
        var newX = x ?? 100;
        var newY = y ?? 600;
        page.AddImage(imagePath,
            new Rectangle(newX, newY, width.HasValue ? newX + width.Value : newX + 200,
                height.HasValue ? newY + height.Value : newY + 200));
        document.Save(outputPath);
        return await Task.FromResult(
            $"Successfully edited image {imageIndex} on page {pageIndex}. Output: {outputPath}");
    }

    /// <summary>
    ///     Extracts images from a PDF
    /// </summary>
    /// <param name="arguments">JSON arguments containing path, outputDirectory, optional pageIndex</param>
    /// <returns>Success message with extracted image count</returns>
    private async Task<string> ExtractImages(JsonObject? arguments)
    {
        var path = ArgumentHelper.GetAndValidatePath(arguments);
        var outputPath = ArgumentHelper.GetStringNullable(arguments, "outputPath");
        var outputDir = ArgumentHelper.GetStringNullable(arguments, "outputDir");
        var pageIndex = ArgumentHelper.GetInt(arguments, "pageIndex");
        var imageIndex = ArgumentHelper.GetIntNullable(arguments, "imageIndex");

        SecurityHelper.ValidateFilePath(path);
        if (!string.IsNullOrEmpty(outputPath))
            SecurityHelper.ValidateFilePath(outputPath, "outputPath");
        if (!string.IsNullOrEmpty(outputDir))
            SecurityHelper.ValidateFilePath(outputDir, "outputDir");

        var targetDir = outputDir ?? Path.GetDirectoryName(outputPath) ?? Path.GetDirectoryName(path) ?? ".";
        Directory.CreateDirectory(targetDir);

        using var document = new Document(path);
        if (pageIndex < 1 || pageIndex > document.Pages.Count)
            throw new ArgumentException($"pageIndex must be between 1 and {document.Pages.Count}");

        var page = document.Pages[pageIndex];
        var images = page.Resources?.Images;
        if (images == null || images.Count == 0)
            return await Task.FromResult($"No images found on page {pageIndex}.");

        if (imageIndex.HasValue)
        {
            if (imageIndex.Value < 1 || imageIndex.Value > images.Count)
                throw new ArgumentException($"imageIndex must be between 1 and {images.Count}");

            var image = images[imageIndex.Value];
            var fileName = outputPath ?? Path.Combine(targetDir, $"page_{pageIndex}_image_{imageIndex.Value}.png");
            await using var imageStream = new FileStream(fileName, FileMode.Create);
#pragma warning disable CA1416 // Validate platform compatibility
            image.Save(imageStream, ImageFormat.Png);
#pragma warning restore CA1416
            return await Task.FromResult($"Extracted image {imageIndex.Value} from page {pageIndex} to: {fileName}");
        }

        var count = 0;
        for (var i = 1; i <= images.Count; i++)
        {
            var image = images[i];
            var fileName = Path.Combine(targetDir, $"page_{pageIndex}_image_{i}.png");
            await using var imageStream = new FileStream(fileName, FileMode.Create);
#pragma warning disable CA1416 // Validate platform compatibility
            image.Save(imageStream, ImageFormat.Png);
#pragma warning restore CA1416
            count++;
        }

        return await Task.FromResult($"Extracted {count} image(s) from page {pageIndex} to: {targetDir}");
    }

    /// <summary>
    ///     Gets all images from a PDF page
    /// </summary>
    /// <param name="arguments">JSON arguments containing path, pageIndex</param>
    /// <returns>Formatted string with all images</returns>
    private async Task<string> GetImages(JsonObject? arguments)
    {
        var path = ArgumentHelper.GetAndValidatePath(arguments);
        var pageIndex = ArgumentHelper.GetInt(arguments, "pageIndex");

        SecurityHelper.ValidateFilePath(path);

        using var document = new Document(path);
        if (pageIndex < 1 || pageIndex > document.Pages.Count)
            throw new ArgumentException($"pageIndex must be between 1 and {document.Pages.Count}");

        var page = document.Pages[pageIndex];
        var images = page.Resources?.Images;

        var sb = new StringBuilder();
        sb.AppendLine("=== PDF Images ===");
        sb.AppendLine();

        if (images == null || images.Count == 0)
        {
            sb.AppendLine($"No images found on page {pageIndex}.");
            return await Task.FromResult(sb.ToString());
        }

        sb.AppendLine($"Page {pageIndex} Images ({images.Count}):");
        sb.AppendLine();

        for (var i = 1; i <= images.Count; i++)
            try
            {
                var image = images[i];
                sb.AppendLine($"[{i}] Image Index: {i}");
                try
                {
                    if (image.Width > 0 && image.Height > 0)
                        sb.AppendLine($"    Size: {image.Width} x {image.Height} pixels");
                }
                catch
                {
                    sb.AppendLine("    Size: (unavailable)");
                }

                sb.AppendLine();
            }
            catch (Exception ex)
            {
                sb.AppendLine($"[{i}] Error reading image: {ex.Message}");
                sb.AppendLine();
            }

        return await Task.FromResult(sb.ToString());
    }
}