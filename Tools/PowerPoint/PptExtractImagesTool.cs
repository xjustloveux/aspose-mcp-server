using System.Text.Json.Nodes;
using Aspose.Slides;
using Aspose.Slides.Export;
using System.Drawing.Imaging;
using System.Diagnostics.CodeAnalysis;

namespace AsposeMcpServer.Tools;

public class PptExtractImagesTool : IAsposeTool
{
    public string Description => "Extract images from PowerPoint slides to a folder";

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
                description = "Output directory (optional, default: same directory as input)"
            },
            format = new
            {
                type = "string",
                description = "Image format: png|jpeg (default: png)"
            }
        },
        required = new[] { "path" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var outputDir = arguments?["outputDir"]?.GetValue<string>() ?? Path.GetDirectoryName(path) ?? ".";
        var formatStr = arguments?["format"]?.GetValue<string>() ?? "png";

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
        var count = 0;

        for (int i = 0; i < presentation.Slides.Count; i++)
        {
            var slide = presentation.Slides[i];
            for (int j = 0; j < slide.Shapes.Count; j++)
            {
                if (slide.Shapes[j] is PictureFrame pic && pic.PictureFormat?.Picture?.Image != null)
                {
                    var image = pic.PictureFormat.Picture.Image;
                    var fileName = Path.Combine(outputDir, $"slide{i + 1}_img{++count}.{extension}");
                    var systemImage = image.SystemImage;
#pragma warning disable CA1416 // Validate platform compatibility
                    systemImage.Save(fileName, format);
#pragma warning restore CA1416 // Validate platform compatibility
                }
            }
        }

        return await Task.FromResult($"已匯出圖片 {count} 張到: {Path.GetFullPath(outputDir)}");
    }
}

