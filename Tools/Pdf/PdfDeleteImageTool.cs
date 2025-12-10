using System.Text.Json.Nodes;
using Aspose.Pdf;

namespace AsposeMcpServer.Tools;

public class PdfDeleteImageTool : IAsposeTool
{
    public string Description => "Delete an image from PDF page (removes image from resources)";

    public object InputSchema => new
    {
        type = "object",
        properties = new
        {
            path = new
            {
                type = "string",
                description = "PDF file path"
            },
            pageIndex = new
            {
                type = "number",
                description = "Page index (1-based)"
            },
            imageIndex = new
            {
                type = "number",
                description = "Image index (0-based, use pdf_extract_images to find index)"
            }
        },
        required = new[] { "path", "pageIndex", "imageIndex" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var pageIndex = arguments?["pageIndex"]?.GetValue<int>() ?? throw new ArgumentException("pageIndex is required");
        var imageIndex = arguments?["imageIndex"]?.GetValue<int>() ?? throw new ArgumentException("imageIndex is required");

        using var document = new Document(path);
        if (pageIndex < 1 || pageIndex > document.Pages.Count)
        {
            throw new ArgumentException($"pageIndex must be between 1 and {document.Pages.Count}");
        }

        var page = document.Pages[pageIndex];
        var images = page.Resources?.Images;
        
        if (images == null || imageIndex < 0 || imageIndex >= images.Count)
        {
            throw new ArgumentException($"imageIndex must be between 0 and {images?.Count - 1 ?? -1}");
        }

        // Note: Deleting images from resources may not remove them from page content
        // Full implementation would require content stream manipulation
        images.Delete(imageIndex);
        document.Save(path);

        return await Task.FromResult($"Image removed from resources (may still appear in content): {path}");
    }
}

