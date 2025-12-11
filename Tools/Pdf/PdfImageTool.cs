using System.Text;
using System.Text.Json.Nodes;
using Aspose.Pdf;
using Aspose.Pdf.Devices;
using AsposeMcpServer.Core;

namespace AsposeMcpServer.Tools;

public class PdfImageTool : IAsposeTool
{
    public string Description => "Manage images in PDF documents (add, delete, edit, extract)";

    public object InputSchema => new
    {
        type = "object",
        properties = new
        {
            operation = new
            {
                type = "string",
                description = "Operation: add, delete, edit, extract",
                @enum = new[] { "add", "delete", "edit", "extract" }
            },
            path = new
            {
                type = "string",
                description = "PDF file path"
            },
            outputPath = new
            {
                type = "string",
                description = "Output file path (optional, defaults to overwrite input for add/delete/edit, required for extract)"
            },
            pageIndex = new
            {
                type = "number",
                description = "Page index (1-based, required for add, delete, edit, extract)"
            },
            imagePath = new
            {
                type = "string",
                description = "Image file path (required for add, edit)"
            },
            imageIndex = new
            {
                type = "number",
                description = "Image index (0-based, required for delete, edit, extract)"
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

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var operation = arguments?["operation"]?.GetValue<string>() ?? throw new ArgumentException("operation is required");

        return operation.ToLower() switch
        {
            "add" => await AddImage(arguments),
            "delete" => await DeleteImage(arguments),
            "edit" => await EditImage(arguments),
            "extract" => await ExtractImages(arguments),
            _ => throw new ArgumentException($"Unknown operation: {operation}")
        };
    }

    private async Task<string> AddImage(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var outputPath = arguments?["outputPath"]?.GetValue<string>() ?? path;
        var pageIndex = arguments?["pageIndex"]?.GetValue<int>() ?? throw new ArgumentException("pageIndex is required");
        var imagePath = arguments?["imagePath"]?.GetValue<string>() ?? throw new ArgumentException("imagePath is required");
        var x = arguments?["x"]?.GetValue<double>() ?? 100;
        var y = arguments?["y"]?.GetValue<double>() ?? 600;
        var width = arguments?["width"]?.GetValue<double?>();
        var height = arguments?["height"]?.GetValue<double?>();

        SecurityHelper.ValidateFilePath(path, "path");
        SecurityHelper.ValidateFilePath(outputPath, "outputPath");
        SecurityHelper.ValidateFilePath(imagePath, "imagePath");

        if (!File.Exists(imagePath))
            throw new FileNotFoundException($"Image file not found: {imagePath}");

        using var document = new Document(path);
        if (pageIndex < 1 || pageIndex > document.Pages.Count)
            throw new ArgumentException($"pageIndex must be between 1 and {document.Pages.Count}");

        var page = document.Pages[pageIndex];
        page.AddImage(imagePath, new Aspose.Pdf.Rectangle(x, y, width.HasValue ? x + width.Value : x + 200, height.HasValue ? y + height.Value : y + 200));
        document.Save(outputPath);
        return await Task.FromResult($"Successfully added image to page {pageIndex}. Output: {outputPath}");
    }

    private async Task<string> DeleteImage(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var outputPath = arguments?["outputPath"]?.GetValue<string>() ?? path;
        var pageIndex = arguments?["pageIndex"]?.GetValue<int>() ?? throw new ArgumentException("pageIndex is required");
        var imageIndex = arguments?["imageIndex"]?.GetValue<int>() ?? throw new ArgumentException("imageIndex is required");

        SecurityHelper.ValidateFilePath(path, "path");
        SecurityHelper.ValidateFilePath(outputPath, "outputPath");

        using var document = new Document(path);
        if (pageIndex < 1 || pageIndex > document.Pages.Count)
            throw new ArgumentException($"pageIndex must be between 1 and {document.Pages.Count}");

        var page = document.Pages[pageIndex];
        var images = page.Resources?.Images;
        if (images == null || imageIndex < 0 || imageIndex >= images!.Count)
            throw new ArgumentException($"imageIndex must be between 0 and {images!.Count - 1}");

        images.Delete(imageIndex);
        document.Save(outputPath);
        return await Task.FromResult($"Successfully deleted image {imageIndex} from page {pageIndex}. Output: {outputPath}");
    }

    private async Task<string> EditImage(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var outputPath = arguments?["outputPath"]?.GetValue<string>() ?? path;
        var pageIndex = arguments?["pageIndex"]?.GetValue<int>() ?? throw new ArgumentException("pageIndex is required");
        var imageIndex = arguments?["imageIndex"]?.GetValue<int>() ?? throw new ArgumentException("imageIndex is required");
        var imagePath = arguments?["imagePath"]?.GetValue<string>() ?? throw new ArgumentException("imagePath is required");
        var x = arguments?["x"]?.GetValue<double?>();
        var y = arguments?["y"]?.GetValue<double?>();
        var width = arguments?["width"]?.GetValue<double?>();
        var height = arguments?["height"]?.GetValue<double?>();

        SecurityHelper.ValidateFilePath(path, "path");
        SecurityHelper.ValidateFilePath(outputPath, "outputPath");
        SecurityHelper.ValidateFilePath(imagePath, "imagePath");

        if (!File.Exists(imagePath))
            throw new FileNotFoundException($"Image file not found: {imagePath}");

        using var document = new Document(path);
        if (pageIndex < 1 || pageIndex > document.Pages.Count)
            throw new ArgumentException($"pageIndex must be between 1 and {document.Pages.Count}");

        var page = document.Pages[pageIndex];
        var images = page.Resources?.Images;
        if (images == null || imageIndex < 0 || imageIndex >= images!.Count)
            throw new ArgumentException($"imageIndex must be between 0 and {images!.Count - 1}");

        images.Delete(imageIndex);
        var newX = x ?? 100;
        var newY = y ?? 600;
        page.AddImage(imagePath, new Aspose.Pdf.Rectangle(newX, newY, width.HasValue ? newX + width.Value : newX + 200, height.HasValue ? newY + height.Value : newY + 200));
        document.Save(outputPath);
        return await Task.FromResult($"Successfully edited image {imageIndex} on page {pageIndex}. Output: {outputPath}");
    }

    private async Task<string> ExtractImages(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var outputPath = arguments?["outputPath"]?.GetValue<string>();
        var outputDir = arguments?["outputDir"]?.GetValue<string>();
        var pageIndex = arguments?["pageIndex"]?.GetValue<int>() ?? throw new ArgumentException("pageIndex is required");
        var imageIndex = arguments?["imageIndex"]?.GetValue<int?>();

        SecurityHelper.ValidateFilePath(path, "path");
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
            if (imageIndex.Value < 0 || imageIndex.Value >= images!.Count)
                throw new ArgumentException($"imageIndex must be between 0 and {images!.Count - 1}");

            var image = images[imageIndex.Value];
            var fileName = outputPath ?? Path.Combine(targetDir, $"page_{pageIndex}_image_{imageIndex.Value}.png");
            using var imageStream = new FileStream(fileName, FileMode.Create);
#pragma warning disable CA1416 // Validate platform compatibility
            image.Save(imageStream, System.Drawing.Imaging.ImageFormat.Png);
#pragma warning restore CA1416
            return await Task.FromResult($"Extracted image {imageIndex.Value} from page {pageIndex} to: {fileName}");
        }
        else
        {
            int count = 0;
            for (int i = 0; i < images.Count; i++)
            {
                var image = images[i];
                var fileName = Path.Combine(targetDir, $"page_{pageIndex}_image_{i}.png");
                using var imageStream = new FileStream(fileName, FileMode.Create);
#pragma warning disable CA1416 // Validate platform compatibility
                image.Save(imageStream, System.Drawing.Imaging.ImageFormat.Png);
#pragma warning restore CA1416
                count++;
            }
            return await Task.FromResult($"Extracted {count} image(s) from page {pageIndex} to: {targetDir}");
        }
    }
}

