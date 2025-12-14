using System.Text.Json.Nodes;
using Aspose.Words;
using AsposeMcpServer.Core;

namespace AsposeMcpServer.Tools;

public class WordWatermarkTool : IAsposeTool
{
    public string Description => @"Manage watermarks in Word documents. Supports 1 operation: add.

Usage examples:
- Add watermark: word_watermark(operation='add', path='doc.docx', text='CONFIDENTIAL', fontSize=72, isSemitransparent=true)";

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
                description = "Document file path (required for all operations)"
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
            fontFamily = new
            {
                type = "string",
                description = "Font family (optional, default: 'Arial')"
            },
            fontSize = new
            {
                type = "number",
                description = "Font size (optional, default: 72)"
            },
            isSemitransparent = new
            {
                type = "boolean",
                description = "Is semitransparent (optional, default: true)"
            },
            layout = new
            {
                type = "string",
                description = "Layout: Diagonal, Horizontal (optional, default: Diagonal)",
                @enum = new[] { "Diagonal", "Horizontal" }
            }
        },
        required = new[] { "operation", "path", "text" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var operation = ArgumentHelper.GetString(arguments, "operation", "operation");

        return operation.ToLower() switch
        {
            "add" => await AddWatermark(arguments),
            _ => throw new ArgumentException($"Unknown operation: {operation}")
        };
    }

    /// <summary>
    /// Adds a watermark to the document
    /// </summary>
    /// <param name="arguments">JSON arguments containing path, optional text, imagePath, outputPath</param>
    /// <returns>Success message</returns>
    private async Task<string> AddWatermark(JsonObject? arguments)
    {
        var path = ArgumentHelper.GetAndValidatePath(arguments);
        var outputPath = ArgumentHelper.GetAndValidateOutputPath(arguments, path);
        SecurityHelper.ValidateFilePath(outputPath, "outputPath");
        var text = ArgumentHelper.GetString(arguments, "text", "text");
        var fontFamily = arguments?["fontFamily"]?.GetValue<string>() ?? "Arial";
        var fontSize = arguments?["fontSize"]?.GetValue<double>() ?? 72;
        var isSemitransparent = arguments?["isSemitransparent"]?.GetValue<bool?>() ?? true;
        var layout = arguments?["layout"]?.GetValue<string>() ?? "Diagonal";

        var doc = new Document(path);
        
        var watermarkOptions = new TextWatermarkOptions
        {
            FontFamily = fontFamily,
            FontSize = (float)fontSize,
            IsSemitrasparent = isSemitransparent,
            Layout = layout.ToLower() == "horizontal" ? WatermarkLayout.Horizontal : WatermarkLayout.Diagonal
        };

        doc.Watermark.SetText(text, watermarkOptions);
        doc.Save(outputPath);

        return await Task.FromResult($"Watermark added to document. Output: {outputPath}");
    }
}

