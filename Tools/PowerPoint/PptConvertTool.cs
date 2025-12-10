using System.Text.Json.Nodes;
using Aspose.Slides;
using Aspose.Slides.Export;

namespace AsposeMcpServer.Tools;

public class PptConvertTool : IAsposeTool
{
    public string Description => "Convert PowerPoint presentation to another format (PDF, HTML, images, etc.)";

    public object InputSchema => new
    {
        type = "object",
        properties = new
        {
            inputPath = new
            {
                type = "string",
                description = "Input file path"
            },
            outputPath = new
            {
                type = "string",
                description = "Output file path"
            },
            format = new
            {
                type = "string",
                description = "Output format (pdf, html, pptx, jpg, png, etc.)"
            }
        },
        required = new[] { "inputPath", "outputPath", "format" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var inputPath = arguments?["inputPath"]?.GetValue<string>() ?? throw new ArgumentException("inputPath is required");
        var outputPath = arguments?["outputPath"]?.GetValue<string>() ?? throw new ArgumentException("outputPath is required");
        var format = arguments?["format"]?.GetValue<string>()?.ToLower() ?? throw new ArgumentException("format is required");

        using var presentation = new Presentation(inputPath);

        SaveFormat saveFormat;
        
        // Handle image formats separately
        // Note: System.Drawing APIs are primarily supported on Windows
        if (format == "jpg" || format == "jpeg")
        {
            // For JPEG, save first slide as image
#pragma warning disable CA1416 // Validate platform compatibility
            using var bitmap = presentation.Slides[0].GetThumbnail(new System.Drawing.Size(1920, 1080));
            bitmap.Save(outputPath, System.Drawing.Imaging.ImageFormat.Jpeg);
#pragma warning restore CA1416
            return await Task.FromResult($"Presentation converted to JPEG: {outputPath}");
        }
        else if (format == "png")
        {
            // For PNG, save first slide as image
#pragma warning disable CA1416 // Validate platform compatibility
            using var bitmap = presentation.Slides[0].GetThumbnail(new System.Drawing.Size(1920, 1080));
            bitmap.Save(outputPath, System.Drawing.Imaging.ImageFormat.Png);
#pragma warning restore CA1416
            return await Task.FromResult($"Presentation converted to PNG: {outputPath}");
        }
        
        saveFormat = format switch
        {
            "pdf" => SaveFormat.Pdf,
            "html" => SaveFormat.Html,
            "pptx" => SaveFormat.Pptx,
            "ppt" => SaveFormat.Ppt,
            "odp" => SaveFormat.Odp,
            "xps" => SaveFormat.Xps,
            "tiff" => SaveFormat.Tiff,
            _ => throw new ArgumentException($"Unsupported format: {format}")
        };

        presentation.Save(outputPath, saveFormat);

        return await Task.FromResult($"Presentation converted from {inputPath} to {outputPath} ({format})");
    }
}

