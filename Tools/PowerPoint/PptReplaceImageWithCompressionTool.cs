using System.Text.Json.Nodes;
using Aspose.Slides;
using Aspose.Slides.Export;
using System.Drawing;
using System.Drawing.Imaging;
using System.Diagnostics.CodeAnalysis;

namespace AsposeMcpServer.Tools;

public class PptReplaceImageWithCompressionTool : IAsposeTool
{
    public string Description => "Replace a picture frame image with optional JPEG compression";

    public object InputSchema => new
    {
        type = "object",
        properties = new
        {
            path = new { type = "string", description = "Presentation file path" },
            slideIndex = new { type = "number", description = "Slide index (0-based)" },
            shapeIndex = new { type = "number", description = "PictureFrame shape index (0-based)" },
            imagePath = new { type = "string", description = "New image file path" },
            jpegQuality = new { type = "number", description = "JPEG quality 10-100 (optional; re-encode as JPEG if provided)" }
        },
        required = new[] { "path", "slideIndex", "shapeIndex", "imagePath" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var slideIndex = arguments?["slideIndex"]?.GetValue<int>() ?? throw new ArgumentException("slideIndex is required");
        var shapeIndex = arguments?["shapeIndex"]?.GetValue<int>() ?? throw new ArgumentException("shapeIndex is required");
        var imagePath = arguments?["imagePath"]?.GetValue<string>() ?? throw new ArgumentException("imagePath is required");
        var jpegQuality = arguments?["jpegQuality"]?.GetValue<int?>();

        using var presentation = new Presentation(path);
        if (slideIndex < 0 || slideIndex >= presentation.Slides.Count)
        {
            throw new ArgumentException($"slideIndex must be between 0 and {presentation.Slides.Count - 1}");
        }

        var slide = presentation.Slides[slideIndex];
        if (shapeIndex < 0 || shapeIndex >= slide.Shapes.Count)
        {
            throw new ArgumentException($"shapeIndex must be between 0 and {slide.Shapes.Count - 1}");
        }

        if (slide.Shapes[shapeIndex] is not PictureFrame pic)
        {
            throw new ArgumentException("指定的 shape 不是圖片框 (PictureFrame)");
        }

        byte[] imageBytes;
        if (jpegQuality.HasValue)
        {
            var quality = Math.Clamp(jpegQuality.Value, 10, 100);
#pragma warning disable CA1416 // Validate platform compatibility
            using var src = Image.FromFile(imagePath);
            var encoder = ImageCodecInfo.GetImageEncoders().First(c => c.FormatID == ImageFormat.Jpeg.Guid);
            var encParams = new EncoderParameters(1);
            encParams.Param[0] = new EncoderParameter(Encoder.Quality, quality);
            using var ms = new MemoryStream();
            src.Save(ms, encoder, encParams);
#pragma warning restore CA1416 // Validate platform compatibility
            imageBytes = ms.ToArray();
        }
        else
        {
            imageBytes = File.ReadAllBytes(imagePath);
        }

        var newImage = presentation.Images.AddImage(imageBytes);
        pic.PictureFormat.Picture.Image = newImage;

        presentation.Save(path, SaveFormat.Pptx);
        return await Task.FromResult($"已替換圖片並套用壓縮 (quality={jpegQuality?.ToString() ?? "原始"})");
    }
}

