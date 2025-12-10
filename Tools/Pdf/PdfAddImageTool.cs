using System.Text.Json.Nodes;
using Aspose.Pdf;

namespace AsposeMcpServer.Tools;

public class PdfAddImageTool : IAsposeTool
{
    public string Description => "Add an image to a PDF document";

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
            imagePath = new
            {
                type = "string",
                description = "Image file path"
            },
            x = new
            {
                type = "number",
                description = "X position (optional, default: 100)"
            },
            y = new
            {
                type = "number",
                description = "Y position (optional, default: 600)"
            },
            width = new
            {
                type = "number",
                description = "Image width (optional)"
            },
            height = new
            {
                type = "number",
                description = "Image height (optional)"
            }
        },
        required = new[] { "path", "pageIndex", "imagePath" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var pageIndex = arguments?["pageIndex"]?.GetValue<int>() ?? throw new ArgumentException("pageIndex is required");
        var imagePath = arguments?["imagePath"]?.GetValue<string>() ?? throw new ArgumentException("imagePath is required");
        var x = arguments?["x"]?.GetValue<double>() ?? 100;
        var y = arguments?["y"]?.GetValue<double>() ?? 600;
        var width = arguments?["width"]?.GetValue<double>();
        var height = arguments?["height"]?.GetValue<double>();

        using var document = new Document(path);
        var page = document.Pages[pageIndex];

        using var imageStream = File.OpenRead(imagePath);
        
        if (width.HasValue && height.HasValue)
        {
            page.AddImage(imageStream, new Rectangle(x, y, x + width.Value, y + height.Value));
        }
        else
        {
            page.AddImage(imageStream, new Rectangle(x, y, x + 200, y + 200));
        }

        document.Save(path);

        return await Task.FromResult($"Image added to page {pageIndex}: {path}");
    }
}

