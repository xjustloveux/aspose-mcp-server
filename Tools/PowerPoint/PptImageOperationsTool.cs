using System.Text.Json.Nodes;
using System.Drawing.Imaging;
using System.Diagnostics.CodeAnalysis;
using Aspose.Slides;
using Aspose.Slides.Export;
using AsposeMcpServer.Core;

namespace AsposeMcpServer.Tools;

/// <summary>
/// Unified tool for PowerPoint image operations (export slides as images, extract images, replace image with compression)
/// Merges: PptExportSlidesAsImagesTool, PptExtractImagesTool, PptReplaceImageWithCompressionTool
/// </summary>
public class PptImageOperationsTool : IAsposeTool
{
    public string Description => "PowerPoint image operations: export slides as images, extract images, or replace image with compression";

    public object InputSchema => new
    {
        type = "object",
        properties = new
        {
            operation = new
            {
                type = "string",
                description = "Operation to perform: 'export_slides', 'extract_images', 'replace_with_compression'",
                @enum = new[] { "export_slides", "extract_images", "replace_with_compression" }
            },
            path = new
            {
                type = "string",
                description = "Presentation file path"
            },
            outputDir = new
            {
                type = "string",
                description = "Output directory (required for export_slides/extract_images)"
            },
            format = new
            {
                type = "string",
                description = "Image format: png|jpeg (optional, default: png)"
            },
            scale = new
            {
                type = "number",
                description = "Scaling factor (optional, for export_slides, default: 1.0)"
            },
            slideIndex = new
            {
                type = "number",
                description = "Slide index (0-based, required for replace_with_compression)"
            },
            shapeIndex = new
            {
                type = "number",
                description = "PictureFrame shape index (0-based, required for replace_with_compression)"
            },
            imagePath = new
            {
                type = "string",
                description = "New image file path (required for replace_with_compression)"
            },
            jpegQuality = new
            {
                type = "number",
                description = "JPEG quality 10-100 (optional, for replace_with_compression, re-encode as JPEG if provided)"
            }
        },
        required = new[] { "operation", "path" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var operation = arguments?["operation"]?.GetValue<string>() ?? throw new ArgumentException("operation is required");
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");

        return operation.ToLower() switch
        {
            "export_slides" => await ExportSlidesAsImagesAsync(arguments, path),
            "extract_images" => await ExtractImagesAsync(arguments, path),
            "replace_with_compression" => await ReplaceImageWithCompressionAsync(arguments, path),
            _ => throw new ArgumentException($"Unknown operation: {operation}")
        };
    }

    private async Task<string> ExportSlidesAsImagesAsync(JsonObject? arguments, string path)
    {
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

    private async Task<string> ExtractImagesAsync(JsonObject? arguments, string path)
    {
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

    private async Task<string> ReplaceImageWithCompressionAsync(JsonObject? arguments, string path)
    {
        var slideIndex = arguments?["slideIndex"]?.GetValue<int>() ?? throw new ArgumentException("slideIndex is required for replace_with_compression operation");
        var shapeIndex = arguments?["shapeIndex"]?.GetValue<int>() ?? throw new ArgumentException("shapeIndex is required for replace_with_compression operation");
        var imagePath = arguments?["imagePath"]?.GetValue<string>() ?? throw new ArgumentException("imagePath is required for replace_with_compression operation");
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
            using var src = System.Drawing.Image.FromFile(imagePath);
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

