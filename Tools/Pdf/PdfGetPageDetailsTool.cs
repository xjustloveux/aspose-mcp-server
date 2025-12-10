using System.Text.Json.Nodes;
using System.Text;
using Aspose.Pdf;

namespace AsposeMcpServer.Tools;

public class PdfGetPageDetailsTool : IAsposeTool
{
    public string Description => "Get detailed information about a PDF page";

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
            }
        },
        required = new[] { "path", "pageIndex" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var pageIndex = arguments?["pageIndex"]?.GetValue<int>() ?? throw new ArgumentException("pageIndex is required");

        using var document = new Document(path);
        if (pageIndex < 1 || pageIndex > document.Pages.Count)
        {
            throw new ArgumentException($"pageIndex must be between 1 and {document.Pages.Count}");
        }

        var page = document.Pages[pageIndex];
        var sb = new StringBuilder();

        sb.AppendLine($"=== Page {pageIndex} Details ===");
        sb.AppendLine($"Size: {page.Rect.Width} x {page.Rect.Height}");
        sb.AppendLine($"Rotation: {page.Rotate}°");
        sb.AppendLine($"MediaBox: ({page.MediaBox.LLX}, {page.MediaBox.LLY}) to ({page.MediaBox.URX}, {page.MediaBox.URY})");
        sb.AppendLine($"CropBox: ({page.CropBox.LLX}, {page.CropBox.LLY}) to ({page.CropBox.URX}, {page.CropBox.URY})");
        sb.AppendLine($"Annotations: {page.Annotations.Count}");
        sb.AppendLine($"Paragraphs: {page.Paragraphs.Count}");
        sb.AppendLine($"Images: {page.Resources?.Images?.Count ?? 0}");

        if (page.Annotations.Count > 0)
        {
            sb.AppendLine("\nAnnotations:");
            for (int i = 0; i < page.Annotations.Count; i++)
            {
                var ann = page.Annotations[i];
                sb.AppendLine($"  [{i}] {ann.GetType().Name}");
            }
        }

        return await Task.FromResult(sb.ToString());
    }
}

