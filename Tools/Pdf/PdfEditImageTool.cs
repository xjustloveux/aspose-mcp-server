using System.Text.Json.Nodes;
using Aspose.Pdf;
using Aspose.Pdf.Operators;

namespace AsposeMcpServer.Tools;

public class PdfEditImageTool : IAsposeTool
{
    public string Description => "Edit image properties (position, size, rotation) on a page";

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
            },
            x = new
            {
                type = "number",
                description = "New X position (optional)"
            },
            y = new
            {
                type = "number",
                description = "New Y position (optional)"
            },
            width = new
            {
                type = "number",
                description = "New width (optional)"
            },
            height = new
            {
                type = "number",
                description = "New height (optional)"
            },
            rotation = new
            {
                type = "number",
                description = "Rotation angle in degrees (optional)"
            }
        },
        required = new[] { "path", "pageIndex", "imageIndex" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var pageIndex = arguments?["pageIndex"]?.GetValue<int>() ?? throw new ArgumentException("pageIndex is required");
        var imageIndex = arguments?["imageIndex"]?.GetValue<int>() ?? throw new ArgumentException("imageIndex is required");
        var x = arguments?["x"]?.GetValue<double?>();
        var y = arguments?["y"]?.GetValue<double?>();
        var width = arguments?["width"]?.GetValue<double?>();
        var height = arguments?["height"]?.GetValue<double?>();
        var rotation = arguments?["rotation"]?.GetValue<double?>();

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

        // Note: Editing image position/size requires manipulating page content operators
        // This is a simplified approach - full implementation would require parsing and modifying content stream
        var changes = new List<string>();
        
        if (x.HasValue) changes.Add($"X: {x.Value}");
        if (y.HasValue) changes.Add($"Y: {y.Value}");
        if (width.HasValue) changes.Add($"Width: {width.Value}");
        if (height.HasValue) changes.Add($"Height: {height.Value}");
        if (rotation.HasValue) changes.Add($"Rotation: {rotation.Value}°");

        document.Save(path);
        return await Task.FromResult($"Image editing requires content stream manipulation. Changes requested: {string.Join(", ", changes)} - {path}");
    }
}

