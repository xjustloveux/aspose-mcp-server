using System.Text.Json.Nodes;
using Aspose.Slides;
using Aspose.Slides.Export;
using System.Drawing.Imaging;
using System.Diagnostics.CodeAnalysis;

namespace AsposeMcpServer.Tools;

public class PptExportSlidesAsImagesTool : IAsposeTool
{
    public string Description => "Export slides as images (PNG/JPEG)";

    public object InputSchema => new
    {
        type = "object",
        properties = new
        {
            path = new
            {
                type = "string",
                description = "Presentation file path"
            },
            outputDir = new
            {
                type = "string",
                description = "Output directory (default: same directory)"
            },
            format = new
            {
                type = "string",
                description = "Image format: png|jpeg (default: png)"
            },
            scale = new
            {
                type = "number",
                description = "Scaling factor (default: 1.0)"
            }
        },
        required = new[] { "path" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var outputDir = arguments?["outputDir"]?.GetValue<string>() ?? Path.GetDirectoryName(path) ?? ".";
        var formatStr = arguments?["format"]?.GetValue<string>() ?? "png";
        var scale = arguments?["scale"]?.GetValue<float?>() ?? 1.0f;

#pragma warning disable CA1416 // Validate platform compatibility
        var format = formatStr.ToLower() switch
        {
            "jpeg" or "jpg" => ImageFormat.Jpeg,
            _ => ImageFormat.Png
        };
        var extension = format == ImageFormat.Png ? "png" : "jpg";
#pragma warning restore CA1416 // Validate platform compatibility

        Directory.CreateDirectory(outputDir);

        using var presentation = new Presentation(path);
        for (int i = 0; i < presentation.Slides.Count; i++)
        {
            var bmp = presentation.Slides[i].GetThumbnail(scale, scale);
            var fileName = Path.Combine(outputDir, $"slide_{i + 1}.{extension}");
#pragma warning disable CA1416 // Validate platform compatibility
            bmp.Save(fileName, format);
#pragma warning restore CA1416 // Validate platform compatibility
        }

        return await Task.FromResult($"已匯出 {presentation.Slides.Count} 張幻燈片到: {Path.GetFullPath(outputDir)}");
    }
}

