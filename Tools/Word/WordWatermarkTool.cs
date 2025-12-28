using System.Text.Json.Nodes;
using Aspose.Words;
using AsposeMcpServer.Core;
using SkiaSharp;

namespace AsposeMcpServer.Tools.Word;

/// <summary>
///     Tool for managing watermarks in Word documents
/// </summary>
public class WordWatermarkTool : IAsposeTool
{
    /// <summary>
    ///     Gets the description of the tool and its usage examples
    /// </summary>
    public string Description => @"Manage watermarks in Word documents. Supports 3 operations: add, add_image, remove.

Usage examples:
- Add text watermark: word_watermark(operation='add', path='doc.docx', text='CONFIDENTIAL', fontSize=72, isSemitransparent=true)
- Add image watermark: word_watermark(operation='add_image', path='doc.docx', imagePath='logo.png', scale=1.0, isWashout=true)
- Remove watermark: word_watermark(operation='remove', path='doc.docx')

Note: On Linux/Docker environments, ensure the specified font (default: Arial) is installed. Missing fonts may cause rendering issues.";

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
- 'add': Add text watermark (required params: path, text)
- 'add_image': Add image watermark (required params: path, imagePath)
- 'remove': Remove watermark from document (required params: path)",
                @enum = new[] { "add", "add_image", "remove" }
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
                description = "Layout: Diagonal, Horizontal (optional, default: Diagonal, for add operation)",
                @enum = new[] { "Diagonal", "Horizontal" }
            },
            // Image watermark parameters
            imagePath = new
            {
                type = "string",
                description =
                    "Image file path for watermark (required for add_image operation). Supports PNG, JPG, BMP, GIF formats."
            },
            scale = new
            {
                type = "number",
                description =
                    "Image scale factor (optional, default: 1.0, for add_image operation). Use 0 for auto-scale to fit page."
            },
            isWashout = new
            {
                type = "boolean",
                description =
                    "Apply washout effect to make image lighter (optional, default: true, for add_image operation)"
            }
        },
        required = new[] { "operation", "path" }
    };

    /// <summary>
    ///     Executes the tool operation with the provided JSON arguments
    /// </summary>
    /// <param name="arguments">JSON arguments object containing operation parameters</param>
    /// <returns>Result message as a string</returns>
    /// <exception cref="ArgumentException">Thrown when operation is unknown or required parameters are missing.</exception>
    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var operation = ArgumentHelper.GetString(arguments, "operation");
        var path = ArgumentHelper.GetAndValidatePath(arguments);
        var outputPath = ArgumentHelper.GetStringNullable(arguments, "outputPath") ?? path;
        SecurityHelper.ValidateFilePath(outputPath, "outputPath", true);

        return operation.ToLower() switch
        {
            "add" => await AddTextWatermarkAsync(path, outputPath, arguments),
            "add_image" => await AddImageWatermarkAsync(path, outputPath, arguments),
            "remove" => await RemoveWatermarkAsync(path, outputPath),
            _ => throw new ArgumentException($"Unknown operation: {operation}")
        };
    }

    /// <summary>
    ///     Adds a text watermark to the document
    /// </summary>
    /// <param name="path">Document file path</param>
    /// <param name="outputPath">Output file path</param>
    /// <param name="arguments">JSON arguments containing text, optional fontFamily, fontSize, isSemitransparent, layout</param>
    /// <returns>Success message with output path</returns>
    private Task<string> AddTextWatermarkAsync(string path, string outputPath, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
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

            return $"Text watermark added to document. Output: {outputPath}";
        });
    }

    /// <summary>
    ///     Adds an image watermark to the document
    /// </summary>
    /// <param name="path">Document file path</param>
    /// <param name="outputPath">Output file path</param>
    /// <param name="arguments">JSON arguments containing imagePath, optional scale, isWashout</param>
    /// <returns>Success message with output path</returns>
    private Task<string> AddImageWatermarkAsync(string path, string outputPath, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var imagePath = ArgumentHelper.GetString(arguments, "imagePath");
            var scale = ArgumentHelper.GetDouble(arguments, "scale", "scale", false, 1.0);
            var isWashout = ArgumentHelper.GetBool(arguments, "isWashout", true);

            SecurityHelper.ValidateFilePath(imagePath, "imagePath", true);

            if (!File.Exists(imagePath))
                throw new FileNotFoundException($"Image file not found: {imagePath}");

            var doc = new Document(path);

            var watermarkOptions = new ImageWatermarkOptions
            {
                Scale = scale,
                IsWashout = isWashout
            };

            using var bitmap = SKBitmap.Decode(imagePath);
            if (bitmap == null)
                throw new ArgumentException(
                    $"Failed to decode image: {imagePath}. Ensure the file is a valid image format.");

            doc.Watermark.SetImage(bitmap, watermarkOptions);
            doc.Save(outputPath);

            return
                $"Image watermark added to document. Image: {Path.GetFileName(imagePath)}, Scale: {scale}, Washout: {isWashout}. Output: {outputPath}";
        });
    }

    /// <summary>
    ///     Removes watermark from the document
    /// </summary>
    /// <param name="path">Document file path</param>
    /// <param name="outputPath">Output file path</param>
    /// <returns>Success message with output path</returns>
    private Task<string> RemoveWatermarkAsync(string path, string outputPath)
    {
        return Task.Run(() =>
        {
            var doc = new Document(path);

            if (doc.Watermark.Type == WatermarkType.None)
                return $"No watermark found in document. Output: {outputPath}";

            doc.Watermark.Remove();
            doc.Save(outputPath);

            return $"Watermark removed from document. Output: {outputPath}";
        });
    }
}