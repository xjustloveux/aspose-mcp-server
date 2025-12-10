using System.Text.Json.Nodes;
using Aspose.Pdf;

namespace AsposeMcpServer.Tools;

public class PdfExtractImagesTool : IAsposeTool
{
    public string Description => "Extract images from a PDF document";

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
            outputDir = new
            {
                type = "string",
                description = "Output directory for extracted images"
            },
            pageIndex = new
            {
                type = "number",
                description = "Page index to extract from (1-based, optional - extracts from all if not specified)"
            }
        },
        required = new[] { "path", "outputDir" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var outputDir = arguments?["outputDir"]?.GetValue<string>() ?? throw new ArgumentException("outputDir is required");
        var pageIndex = arguments?["pageIndex"]?.GetValue<int>();

        Directory.CreateDirectory(outputDir);

        using var document = new Document(path);
        int imageCount = 0;

        var pagesToProcess = pageIndex.HasValue 
            ? new[] { document.Pages[pageIndex.Value] } 
            : document.Pages.Cast<Page>();

        foreach (var page in pagesToProcess)
        {
            foreach (var xImage in page.Resources.Images)
            {
                var outputPath = Path.Combine(outputDir, $"image_{++imageCount}.png");
                using var outputStream = File.Create(outputPath);
                if (OperatingSystem.IsWindows())
                {
                    xImage.Save(outputStream, System.Drawing.Imaging.ImageFormat.Png);
                }
                else
                {
                    // For non-Windows platforms, save as stream
                    xImage.Save(outputStream);
                }
            }
        }

        return await Task.FromResult($"Extracted {imageCount} images to: {outputDir}");
    }
}

