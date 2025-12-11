using System.Text.Json.Nodes;
using Aspose.Pdf;
using Aspose.Pdf.Text;
using AsposeMcpServer.Core;

namespace AsposeMcpServer.Tools;

public class PdfWatermarkTool : IAsposeTool
{
    public string Description => "Manage watermarks in PDF documents (add text watermark)";

    public object InputSchema => new
    {
        type = "object",
        properties = new
        {
            operation = new
            {
                type = "string",
                description = "Operation: add",
                @enum = new[] { "add" }
            },
            path = new
            {
                type = "string",
                description = "PDF file path"
            },
            outputPath = new
            {
                type = "string",
                description = "Output file path (optional, defaults to overwrite input)"
            },
            text = new
            {
                type = "string",
                description = "Watermark text (required for add)"
            },
            opacity = new
            {
                type = "number",
                description = "Opacity (0-1, optional, default: 0.3)"
            },
            fontSize = new
            {
                type = "number",
                description = "Font size (optional, default: 72)"
            },
            fontName = new
            {
                type = "string",
                description = "Font name (optional, default: 'Arial')"
            },
            rotation = new
            {
                type = "number",
                description = "Rotation angle in degrees (optional, default: 45)"
            },
            horizontalAlignment = new
            {
                type = "string",
                description = "Horizontal alignment: Left, Center, Right (optional, default: Center)",
                @enum = new[] { "Left", "Center", "Right" }
            },
            verticalAlignment = new
            {
                type = "string",
                description = "Vertical alignment: Top, Center, Bottom (optional, default: Center)",
                @enum = new[] { "Top", "Center", "Bottom" }
            }
        },
        required = new[] { "operation", "path", "text" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var operation = arguments?["operation"]?.GetValue<string>() ?? throw new ArgumentException("operation is required");

        return operation.ToLower() switch
        {
            "add" => await AddWatermark(arguments),
            _ => throw new ArgumentException($"Unknown operation: {operation}")
        };
    }

    private async Task<string> AddWatermark(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var outputPath = arguments?["outputPath"]?.GetValue<string>() ?? path;
        var text = arguments?["text"]?.GetValue<string>() ?? throw new ArgumentException("text is required");
        var opacity = arguments?["opacity"]?.GetValue<double>() ?? 0.3;
        var fontSize = arguments?["fontSize"]?.GetValue<double>() ?? 72;
        var fontName = arguments?["fontName"]?.GetValue<string>() ?? "Arial";
        var rotation = arguments?["rotation"]?.GetValue<double>() ?? 45;

        SecurityHelper.ValidateFilePath(path, "path");
        SecurityHelper.ValidateFilePath(outputPath, "outputPath");
        var horizontalAlignment = arguments?["horizontalAlignment"]?.GetValue<string>() ?? "Center";
        var verticalAlignment = arguments?["verticalAlignment"]?.GetValue<string>() ?? "Center";

        using var document = new Document(path);

        var hAlign = horizontalAlignment.ToLower() switch
        {
            "left" => HorizontalAlignment.Left,
            "right" => HorizontalAlignment.Right,
            _ => HorizontalAlignment.Center
        };

        var vAlign = verticalAlignment.ToLower() switch
        {
            "top" => VerticalAlignment.Top,
            "bottom" => VerticalAlignment.Bottom,
            _ => VerticalAlignment.Center
        };

        foreach (Page page in document.Pages)
        {
            var watermark = new WatermarkArtifact();
            var textState = new TextState
            {
                FontSize = (float)fontSize,
                ForegroundColor = Aspose.Pdf.Color.Gray,
                Font = FontRepository.FindFont(fontName)
            };

            watermark.SetTextAndState(text, textState);
            watermark.ArtifactHorizontalAlignment = hAlign;
            watermark.ArtifactVerticalAlignment = vAlign;
            watermark.Rotation = rotation;
            watermark.Opacity = opacity;

            page.Artifacts.Add(watermark);
        }

        document.Save(outputPath);

        return await Task.FromResult($"Watermark added to PDF. Output: {outputPath}");
    }
}

