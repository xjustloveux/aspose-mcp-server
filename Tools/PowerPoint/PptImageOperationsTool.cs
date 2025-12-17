using System.Drawing;
using System.Drawing.Imaging;
using System.Text.Json.Nodes;
using Aspose.Slides;
using Aspose.Slides.Export;
using AsposeMcpServer.Core;

namespace AsposeMcpServer.Tools.PowerPoint;

/// <summary>
///     Unified tool for PowerPoint image operations (export slides as images, extract images, replace image with
///     compression)
///     Merges: PptExportSlidesAsImagesTool, PptExtractImagesTool, PptReplaceImageWithCompressionTool
/// </summary>
public class PptImageOperationsTool : IAsposeTool
{
    public string Description =>
        @"PowerPoint image operations. Supports 3 operations: export_slides, extract_images, replace_with_compression.

Usage examples:
- Export slides as images: ppt_image_operations(operation='export_slides', path='presentation.pptx', outputDir='images/', format='png')
- Extract images: ppt_image_operations(operation='extract_images', path='presentation.pptx', outputDir='images/')
- Replace with compression: ppt_image_operations(operation='replace_with_compression', path='presentation.pptx', slideIndex=0, shapeIndex=0, imagePath='new_image.png')";

    public object InputSchema => new
    {
        type = "object",
        properties = new
        {
            operation = new
            {
                type = "string",
                description = @"Operation to perform.
- 'export_slides': Export slides as images (required params: path, outputDir)
- 'extract_images': Extract images from presentation (required params: path, outputDir)
- 'replace_with_compression': Replace image with compression (required params: path, slideIndex, shapeIndex, imagePath)",
                @enum = new[] { "export_slides", "extract_images", "replace_with_compression" }
            },
            path = new
            {
                type = "string",
                description = "Presentation file path (required for all operations)"
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
                description =
                    "JPEG quality 10-100 (optional, for replace_with_compression, re-encode as JPEG if provided)"
            },
            outputPath = new
            {
                type = "string",
                description = "Output file path (optional, for all operations, defaults to input path)"
            }
        },
        required = new[] { "operation", "path" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var operation = ArgumentHelper.GetString(arguments, "operation");
        var path = ArgumentHelper.GetAndValidatePath(arguments);

        return operation.ToLower() switch
        {
            "export_slides" => await ExportSlidesAsImagesAsync(arguments, path),
            "extract_images" => await ExtractImagesAsync(arguments, path),
            "replace_with_compression" => await ReplaceImageWithCompressionAsync(arguments, path),
            _ => throw new ArgumentException($"Unknown operation: {operation}")
        };
    }

    /// <summary>
    ///     Exports slides as images
    /// </summary>
    /// <param name="arguments">JSON arguments containing outputDirectory, optional slideIndexes, format, outputPath</param>
    /// <param name="path">PowerPoint file path</param>
    /// <returns>Success message with exported image count</returns>
    private async Task<string> ExportSlidesAsImagesAsync(JsonObject? arguments, string path)
    {
        var outputDir = ArgumentHelper.GetStringNullable(arguments, "outputDir") ?? Path.GetDirectoryName(path) ?? ".";
        var formatStr = ArgumentHelper.GetString(arguments, "format", "png");
        var scale = ArgumentHelper.GetFloat(arguments, "scale", 1.0f);

#pragma warning disable CA1416 // Validate platform compatibility
        var format = formatStr.ToLower() switch
        {
            "jpeg" or "jpg" => ImageFormat.Jpeg,
            _ => ImageFormat.Png
        };
        var extension = format.Equals(ImageFormat.Png) ? "png" : "jpg";
#pragma warning restore CA1416 // Validate platform compatibility

        Directory.CreateDirectory(outputDir);

        using var presentation = new Presentation(path);
        for (var i = 0; i < presentation.Slides.Count; i++)
        {
            var bmp = presentation.Slides[i].GetThumbnail(scale, scale);
            var fileName = Path.Combine(outputDir, $"slide_{i + 1}.{extension}");
#pragma warning disable CA1416 // Validate platform compatibility
            bmp.Save(fileName, format);
#pragma warning restore CA1416 // Validate platform compatibility
        }

        return await Task.FromResult($"Exported {presentation.Slides.Count} slides to: {Path.GetFullPath(outputDir)}");
    }

    /// <summary>
    ///     Extracts images from the presentation
    /// </summary>
    /// <param name="arguments">JSON arguments containing outputDirectory, optional slideIndex</param>
    /// <param name="path">PowerPoint file path</param>
    /// <returns>Success message with extracted image count</returns>
    private async Task<string> ExtractImagesAsync(JsonObject? arguments, string path)
    {
        var outputDir = ArgumentHelper.GetStringNullable(arguments, "outputDir") ?? Path.GetDirectoryName(path) ?? ".";
        var formatStr = ArgumentHelper.GetString(arguments, "format", "png");

#pragma warning disable CA1416 // Validate platform compatibility
        var format = formatStr.ToLower() switch
        {
            "jpeg" or "jpg" => ImageFormat.Jpeg,
            _ => ImageFormat.Png
        };
        var extension = format.Equals(ImageFormat.Png) ? "png" : "jpg";
#pragma warning restore CA1416 // Validate platform compatibility

        Directory.CreateDirectory(outputDir);

        using var presentation = new Presentation(path);
        var count = 0;

        var slideIndex = 0;
        foreach (var slide in presentation.Slides)
        {
            slideIndex++;
            foreach (var shape in slide.Shapes)
                if (shape is PictureFrame { PictureFormat.Picture.Image: not null } pic)
                {
                    var image = pic.PictureFormat.Picture.Image;
                    var fileName = Path.Combine(outputDir, $"slide{slideIndex}_img{++count}.{extension}");
                    var systemImage = image.SystemImage;
#pragma warning disable CA1416
                    // Validate platform compatibility
                    systemImage.Save(fileName, format);
#pragma warning restore CA1416
                    // Validate platform compatibility
                }
        }

        return await Task.FromResult($"Exported {count} images to: {Path.GetFullPath(outputDir)}");
    }

    /// <summary>
    ///     Replaces image with compression
    /// </summary>
    /// <param name="arguments">
    ///     JSON arguments containing slideIndex, imageIndex, newImagePath, optional compressionLevel,
    ///     outputPath
    /// </param>
    /// <param name="path">PowerPoint file path</param>
    /// <returns>Success message</returns>
    private async Task<string> ReplaceImageWithCompressionAsync(JsonObject? arguments, string path)
    {
        var slideIndex = ArgumentHelper.GetInt(arguments, "slideIndex");
        var shapeIndex = ArgumentHelper.GetInt(arguments, "shapeIndex");
        var imagePath = ArgumentHelper.GetString(arguments, "imagePath");
        var jpegQuality = ArgumentHelper.GetIntNullable(arguments, "jpegQuality");

        using var presentation = new Presentation(path);
        if (slideIndex < 0 || slideIndex >= presentation.Slides.Count)
            throw new ArgumentException($"slideIndex must be between 0 and {presentation.Slides.Count - 1}");

        var slide = presentation.Slides[slideIndex];
        if (shapeIndex < 0 || shapeIndex >= slide.Shapes.Count)
            throw new ArgumentException($"shapeIndex must be between 0 and {slide.Shapes.Count - 1}");

        if (slide.Shapes[shapeIndex] is not PictureFrame pic)
            throw new ArgumentException("The specified shape is not a PictureFrame");

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
            imageBytes = await File.ReadAllBytesAsync(imagePath);
        }

        var newImage = presentation.Images.AddImage(imageBytes);
        pic.PictureFormat.Picture.Image = newImage;

        var outputPath = ArgumentHelper.GetAndValidateOutputPath(arguments, path);
        presentation.Save(outputPath, SaveFormat.Pptx);
        return await Task.FromResult(
            $"Image replaced with compression applied (quality={jpegQuality?.ToString() ?? "original"})");
    }
}