using System.ComponentModel;
using Aspose.Words;
using AsposeMcpServer.Core.Helpers;
using AsposeMcpServer.Core.Session;
using ModelContextProtocol.Server;
using SkiaSharp;

namespace AsposeMcpServer.Tools.Word;

/// <summary>
///     Tool for managing watermarks in Word documents
/// </summary>
[McpServerToolType]
public class WordWatermarkTool
{
    /// <summary>
    ///     Session manager for document session operations
    /// </summary>
    private readonly DocumentSessionManager? _sessionManager;

    /// <summary>
    ///     Initializes a new instance of the WordWatermarkTool class
    /// </summary>
    /// <param name="sessionManager">Optional session manager for in-memory document operations</param>
    public WordWatermarkTool(DocumentSessionManager? sessionManager = null)
    {
        _sessionManager = sessionManager;
    }

    [McpServerTool(Name = "word_watermark")]
    [Description(@"Manage watermarks in Word documents. Supports 3 operations: add, add_image, remove.

Usage examples:
- Add text watermark: word_watermark(operation='add', path='doc.docx', text='CONFIDENTIAL', fontSize=72, isSemitransparent=true)
- Add image watermark: word_watermark(operation='add_image', path='doc.docx', imagePath='logo.png', scale=1.0, isWashout=true)
- Remove watermark: word_watermark(operation='remove', path='doc.docx')

Note: On Linux/Docker environments, ensure the specified font (default: Arial) is installed. Missing fonts may cause rendering issues.")]
    public string Execute(
        [Description("Operation to perform: 'add' (text watermark), 'add_image' (image watermark), 'remove'")]
        string operation,
        [Description("Document file path (required if no sessionId)")]
        string? path = null,
        [Description("Session ID for in-memory editing")]
        string? sessionId = null,
        [Description("Output file path (file mode only, defaults to overwrite input)")]
        string? outputPath = null,
        [Description("Watermark text (required for add)")]
        string? text = null,
        [Description("Font family (optional, default: 'Arial')")]
        string fontFamily = "Arial",
        [Description("Font size (optional, default: 72)")]
        double fontSize = 72,
        [Description("Is semitransparent (optional, default: true)")]
        bool isSemitransparent = true,
        [Description("Layout: 'Diagonal' or 'Horizontal' (optional, default: 'Diagonal')")]
        string layout = "Diagonal",
        [Description("Image file path for watermark (required for add_image). Supports PNG, JPG, BMP, GIF formats.")]
        string? imagePath = null,
        [Description("Image scale factor (optional, default: 1.0). Use 0 for auto-scale to fit page.")]
        double scale = 1.0,
        [Description("Apply washout effect to make image lighter (optional, default: true)")]
        bool isWashout = true)
    {
        using var ctx = DocumentContext<Document>.Create(_sessionManager, sessionId, path);
        var effectiveOutputPath = outputPath ?? path;

        return operation.ToLower() switch
        {
            "add" => AddTextWatermark(ctx, effectiveOutputPath, text, fontFamily, fontSize, isSemitransparent, layout),
            "add_image" => AddImageWatermark(ctx, effectiveOutputPath, imagePath, scale, isWashout),
            "remove" => RemoveWatermark(ctx, effectiveOutputPath),
            _ => throw new ArgumentException($"Unknown operation: {operation}")
        };
    }

    /// <summary>
    ///     Adds a text watermark to the document.
    /// </summary>
    /// <param name="ctx">The document context.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="text">The watermark text.</param>
    /// <param name="fontFamily">The font family for the watermark text.</param>
    /// <param name="fontSize">The font size for the watermark text.</param>
    /// <param name="isSemitransparent">Whether the watermark should be semitransparent.</param>
    /// <param name="layout">The watermark layout (Diagonal or Horizontal).</param>
    /// <returns>A message indicating the result of the operation.</returns>
    /// <exception cref="ArgumentException">Thrown when text is null or empty.</exception>
    private static string AddTextWatermark(DocumentContext<Document> ctx, string? outputPath, string? text,
        string fontFamily, double fontSize, bool isSemitransparent, string layout)
    {
        if (string.IsNullOrEmpty(text))
            throw new ArgumentException("Text is required for add operation");

        var doc = ctx.Document;

        var watermarkOptions = new TextWatermarkOptions
        {
            FontFamily = fontFamily,
            FontSize = (float)fontSize,
            IsSemitrasparent = isSemitransparent,
            Layout = layout.ToLower() == "horizontal" ? WatermarkLayout.Horizontal : WatermarkLayout.Diagonal
        };

        doc.Watermark.SetText(text, watermarkOptions);

        ctx.Save(outputPath);

        return $"Text watermark added to document.\n{ctx.GetOutputMessage(outputPath)}";
    }

    /// <summary>
    ///     Adds an image watermark to the document.
    /// </summary>
    /// <param name="ctx">The document context.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="imagePath">The image file path for the watermark.</param>
    /// <param name="scale">The scale factor for the image.</param>
    /// <param name="isWashout">Whether to apply washout effect to the image.</param>
    /// <returns>A message indicating the result of the operation.</returns>
    /// <exception cref="ArgumentException">Thrown when imagePath is null or empty, or image cannot be decoded.</exception>
    /// <exception cref="FileNotFoundException">Thrown when the image file is not found.</exception>
    private static string AddImageWatermark(DocumentContext<Document> ctx, string? outputPath,
        string? imagePath, double scale, bool isWashout)
    {
        if (string.IsNullOrEmpty(imagePath))
            throw new ArgumentException("imagePath is required for add_image operation");

        SecurityHelper.ValidateFilePath(imagePath, "imagePath", true);

        if (!File.Exists(imagePath))
            throw new FileNotFoundException($"Image file not found: {imagePath}");

        var doc = ctx.Document;

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

        ctx.Save(outputPath);

        return
            $"Image watermark added to document. Image: {Path.GetFileName(imagePath)}, Scale: {scale}, Washout: {isWashout}.\n{ctx.GetOutputMessage(outputPath)}";
    }

    /// <summary>
    ///     Removes watermark from the document.
    /// </summary>
    /// <param name="ctx">The document context.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <returns>A message indicating the result of the operation.</returns>
    private static string RemoveWatermark(DocumentContext<Document> ctx, string? outputPath)
    {
        var doc = ctx.Document;

        if (doc.Watermark.Type == WatermarkType.None)
            return $"No watermark found in document.\n{ctx.GetOutputMessage(outputPath)}";

        doc.Watermark.Remove();

        ctx.Save(outputPath);

        return $"Watermark removed from document.\n{ctx.GetOutputMessage(outputPath)}";
    }
}