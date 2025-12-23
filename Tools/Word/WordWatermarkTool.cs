using System.Text.Json.Nodes;
using Aspose.Words;
using AsposeMcpServer.Core;

namespace AsposeMcpServer.Tools.Word;

/// <summary>
///     Tool for managing watermarks in Word documents
/// </summary>
public class WordWatermarkTool : IAsposeTool
{
    /// <summary>
    ///     Gets the description of the tool and its usage examples
    /// </summary>
    public string Description => @"Manage watermarks in Word documents. Supports 1 operation: add.

Usage examples:
- Add watermark: word_watermark(operation='add', path='doc.docx', text='CONFIDENTIAL', fontSize=72, isSemitransparent=true)";

    /// <summary>
    ///     Gets the JSON schema defining the input parameters for the tool
    /// </summary>
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

    /// <summary>
    ///     Executes the tool operation with the provided JSON arguments
    /// </summary>
    /// <param name="arguments">JSON arguments object containing operation parameters</param>
    /// <returns>Result message as a string</returns>
    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var operation = ArgumentHelper.GetString(arguments, "operation");

        return operation.ToLower() switch
        {
            "add" => await AddWatermark(arguments),
            _ => throw new ArgumentException($"Unknown operation: {operation}")
        };
    }

    /// <summary>
    ///     Adds a watermark to the document
    /// </summary>
    /// <param name="arguments">JSON arguments containing path, optional text, imagePath, outputPath</param>
    /// <returns>Success message</returns>
    private Task<string> AddWatermark(JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var path = ArgumentHelper.GetAndValidatePath(arguments);
            var outputPath = ArgumentHelper.GetAndValidateOutputPath(arguments, path);
            SecurityHelper.ValidateFilePath(outputPath, "outputPath", true);
            var text = ArgumentHelper.GetString(arguments, "text");
            var fontFamily = ArgumentHelper.GetString(arguments, "fontFamily", "Arial");
            var fontSize = ArgumentHelper.GetDouble(arguments, "fontSize", "fontSize", false, 72);
            var isSemitransparent = ArgumentHelper.GetBool(arguments, "isSemitransparent", true);
            var layout = ArgumentHelper.GetString(arguments, "layout", "Diagonal");

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

            return $"Watermark added to document. Output: {outputPath}";
        });
    }
}