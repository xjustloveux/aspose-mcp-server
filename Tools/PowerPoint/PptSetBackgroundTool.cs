using System.Text.Json.Nodes;
using Aspose.Slides;
using Aspose.Slides.Export;
using System.Drawing;

namespace AsposeMcpServer.Tools;

public class PptSetBackgroundTool : IAsposeTool
{
    public string Description => "Set slide background color or image";

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
            slideIndex = new
            {
                type = "number",
                description = "Slide index (0-based, optional, default: 0)"
            },
            color = new
            {
                type = "string",
                description = "Hex color like #FFAA00 (optional)"
            },
            imagePath = new
            {
                type = "string",
                description = "Background image path (optional)"
            }
        },
        required = new[] { "path" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var slideIndex = arguments?["slideIndex"]?.GetValue<int?>() ?? 0;
        var colorHex = arguments?["color"]?.GetValue<string>();
        var imagePath = arguments?["imagePath"]?.GetValue<string>();

        using var presentation = new Presentation(path);
        if (slideIndex < 0 || slideIndex >= presentation.Slides.Count)
        {
            throw new ArgumentException($"slideIndex must be between 0 and {presentation.Slides.Count - 1}");
        }

        var slide = presentation.Slides[slideIndex];
        var fillFormat = slide.Background.FillFormat;

        if (!string.IsNullOrWhiteSpace(imagePath))
        {
            var img = presentation.Images.AddImage(File.ReadAllBytes(imagePath));
            fillFormat.FillType = FillType.Picture;
            fillFormat.PictureFillFormat.PictureFillMode = PictureFillMode.Stretch;
            fillFormat.PictureFillFormat.Picture.Image = img;
        }
        else if (!string.IsNullOrWhiteSpace(colorHex))
        {
            var color = ColorTranslator.FromHtml(colorHex);
            fillFormat.FillType = FillType.Solid;
            fillFormat.SolidFillColor.Color = color;
        }
        else
        {
            throw new ArgumentException("請至少提供 color 或 imagePath 其中之一");
        }

        presentation.Save(path, SaveFormat.Pptx);
        return await Task.FromResult($"已更新投影片 {slideIndex} 背景: {path}");
    }
}

