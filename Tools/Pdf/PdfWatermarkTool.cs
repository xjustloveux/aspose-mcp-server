using System.Text.Json.Nodes;
using Aspose.Pdf;
using Aspose.Pdf.Text;
using AsposeMcpServer.Core;

namespace AsposeMcpServer.Tools.Pdf;

/// <summary>
///     Tool for managing watermarks in PDF documents
/// </summary>
public class PdfWatermarkTool : IAsposeTool
{
    public string Description => @"Manage watermarks in PDF documents. Supports 1 operation: add.

Usage examples:
- Add watermark: pdf_watermark(operation='add', path='doc.pdf', text='CONFIDENTIAL', fontSize=72, opacity=0.3)";

    public object InputSchema => new
    {
        type = "object",
        properties = new
        {
            operation = new
            {
                type = "string",
                description = @"Operation to perform.
- 'add': Add text watermark (required params: path, text)",
                @enum = new[] { "add" }
            },
            path = new
            {
                type = "string",
                description = "PDF file path (required for all operations)"
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

    /// <summary>
    ///     Executes the tool operation with the provided JSON arguments
    /// </summary>
    /// <param name="arguments">JSON arguments object containing operation parameters</param>
    /// <returns>Result message as a string</returns>
    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var operation = ArgumentHelper.GetString(arguments, "operation");
        var path = ArgumentHelper.GetAndValidatePath(arguments);
        var outputPath = ArgumentHelper.GetAndValidateOutputPath(arguments, path);

        return operation.ToLower() switch
        {
            "add" => await AddWatermark(path, outputPath, arguments),
            _ => throw new ArgumentException($"Unknown operation: {operation}")
        };
    }

    /// <summary>
    ///     Adds a watermark to the PDF
    /// </summary>
    /// <param name="path">Input file path</param>
    /// <param name="outputPath">Output file path</param>
    /// <param name="arguments">JSON arguments containing text, optional opacity, fontSize, fontName, rotation, alignment</param>
    /// <returns>Success message</returns>
    private Task<string> AddWatermark(string path, string outputPath, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var text = ArgumentHelper.GetString(arguments, "text");
            var opacity = ArgumentHelper.GetDouble(arguments, "opacity", "opacity", false, 0.3);
            var fontSize = ArgumentHelper.GetDouble(arguments, "fontSize", "fontSize", false, 72);
            var fontName = ArgumentHelper.GetString(arguments, "fontName", "Arial");
            var rotation = ArgumentHelper.GetDouble(arguments, "rotation", "rotation", false, 45);
            var horizontalAlignment = ArgumentHelper.GetString(arguments, "horizontalAlignment", "Center");
            var verticalAlignment = ArgumentHelper.GetString(arguments, "verticalAlignment", "Center");

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

            foreach (var page in document.Pages)
            {
                var watermark = new WatermarkArtifact();
                var textState = new TextState
                {
                    ForegroundColor = Color.Gray
                };

                // Apply font settings using FontHelper
                FontHelper.Pdf.ApplyFontSettings(
                    textState,
                    fontName,
                    fontSize
                );

                watermark.SetTextAndState(text, textState);
                watermark.ArtifactHorizontalAlignment = hAlign;
                watermark.ArtifactVerticalAlignment = vAlign;
                watermark.Rotation = rotation;
                watermark.Opacity = opacity;

                page.Artifacts.Add(watermark);
            }

            document.Save(outputPath);

            return $"Watermark added to PDF. Output: {outputPath}";
        });
    }
}