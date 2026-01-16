using System.ComponentModel;
using Aspose.Words;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Session;
using ModelContextProtocol.Server;

namespace AsposeMcpServer.Tools.Word;

/// <summary>
///     Tool for managing watermarks in Word documents
/// </summary>
[McpServerToolType]
public class WordWatermarkTool
{
    /// <summary>
    ///     Handler registry for watermark operations
    /// </summary>
    private readonly HandlerRegistry<Document> _handlerRegistry;

    /// <summary>
    ///     Identity accessor for session isolation
    /// </summary>
    private readonly ISessionIdentityAccessor? _identityAccessor;

    /// <summary>
    ///     Session manager for document session operations
    /// </summary>
    private readonly DocumentSessionManager? _sessionManager;

    /// <summary>
    ///     Initializes a new instance of the WordWatermarkTool class
    /// </summary>
    /// <param name="sessionManager">Optional session manager for in-memory document operations</param>
    /// <param name="identityAccessor">Optional identity accessor for session isolation</param>
    public WordWatermarkTool(DocumentSessionManager? sessionManager = null,
        ISessionIdentityAccessor? identityAccessor = null)
    {
        _sessionManager = sessionManager;
        _identityAccessor = identityAccessor;
        _handlerRegistry = HandlerRegistry<Document>.CreateFromNamespace("AsposeMcpServer.Handlers.Word.Watermark");
    }

    /// <summary>
    ///     Executes a Word watermark operation (add, add_image, remove).
    /// </summary>
    /// <param name="operation">The operation to perform: add (text watermark), add_image (image watermark), remove.</param>
    /// <param name="path">Word document file path (required if no sessionId).</param>
    /// <param name="sessionId">Session ID for in-memory editing.</param>
    /// <param name="outputPath">Output file path (file mode only, defaults to overwrite input).</param>
    /// <param name="text">Watermark text (required for add).</param>
    /// <param name="fontFamily">Font family (for add, default: Arial).</param>
    /// <param name="fontSize">Font size (for add, default: 72).</param>
    /// <param name="isSemitransparent">Semitransparent watermark (for add, default: true).</param>
    /// <param name="layout">Layout: Diagonal or Horizontal (for add, default: Diagonal).</param>
    /// <param name="imagePath">Image file path for watermark (required for add_image).</param>
    /// <param name="scale">Image scale factor (for add_image, default: 1.0).</param>
    /// <param name="isWashout">Apply washout effect to image (for add_image, default: true).</param>
    /// <returns>A message indicating the result of the operation.</returns>
    /// <exception cref="ArgumentException">Thrown when required parameters are missing or the operation is unknown.</exception>
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
        var parameters = BuildParameters(operation, text, fontFamily, fontSize, isSemitransparent, layout, imagePath,
            scale, isWashout);

        var handler = _handlerRegistry.GetHandler(operation);

        using var ctx = DocumentContext<Document>.Create(_sessionManager, sessionId, path, _identityAccessor);

        var effectiveOutputPath = outputPath ?? path;

        var operationContext = new OperationContext<Document>
        {
            Document = ctx.Document,
            SessionManager = _sessionManager,
            IdentityAccessor = _identityAccessor,
            SessionId = sessionId,
            SourcePath = path,
            OutputPath = effectiveOutputPath
        };

        var result = handler.Execute(operationContext, parameters);

        if (operationContext.IsModified)
            ctx.Save(effectiveOutputPath);

        return ctx.IsSession ? result : $"{result}\n{ctx.GetOutputMessage(effectiveOutputPath)}";
    }

    /// <summary>
    ///     Builds OperationParameters from method parameters using strategy pattern.
    /// </summary>
    private static OperationParameters BuildParameters(
        string operation,
        string? text,
        string fontFamily,
        double fontSize,
        bool isSemitransparent,
        string layout,
        string? imagePath,
        double scale,
        bool isWashout)
    {
        return operation.ToLower() switch
        {
            "add" => BuildAddParameters(text, fontFamily, fontSize, isSemitransparent, layout),
            "add_image" => BuildAddImageParameters(imagePath, scale, isWashout),
            _ => new OperationParameters()
        };
    }

    /// <summary>
    ///     Builds parameters for the add text watermark operation.
    /// </summary>
    /// <param name="text">The watermark text.</param>
    /// <param name="fontFamily">The font family name.</param>
    /// <param name="fontSize">The font size.</param>
    /// <param name="isSemitransparent">Whether the watermark is semitransparent.</param>
    /// <param name="layout">The watermark layout (Diagonal or Horizontal).</param>
    /// <returns>OperationParameters configured for adding a text watermark.</returns>
    private static OperationParameters BuildAddParameters(string? text, string fontFamily, double fontSize,
        bool isSemitransparent, string layout)
    {
        var parameters = new OperationParameters();
        if (text != null) parameters.Set("text", text);
        parameters.Set("fontFamily", fontFamily);
        parameters.Set("fontSize", fontSize);
        parameters.Set("isSemitransparent", isSemitransparent);
        parameters.Set("layout", layout);
        return parameters;
    }

    /// <summary>
    ///     Builds parameters for the add image watermark operation.
    /// </summary>
    /// <param name="imagePath">The path to the watermark image file.</param>
    /// <param name="scale">The image scale factor.</param>
    /// <param name="isWashout">Whether to apply the washout effect.</param>
    /// <returns>OperationParameters configured for adding an image watermark.</returns>
    private static OperationParameters BuildAddImageParameters(string? imagePath, double scale, bool isWashout)
    {
        var parameters = new OperationParameters();
        if (imagePath != null) parameters.Set("imagePath", imagePath);
        parameters.Set("scale", scale);
        parameters.Set("isWashout", isWashout);
        return parameters;
    }
}
