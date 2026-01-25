using System.ComponentModel;
using Aspose.Pdf;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Session;
using AsposeMcpServer.Helpers;
using ModelContextProtocol.Server;

namespace AsposeMcpServer.Tools.Pdf;

/// <summary>
///     Tool for managing watermarks in PDF documents
/// </summary>
[ToolHandlerMapping("AsposeMcpServer.Handlers.Pdf.Watermark")]
[McpServerToolType]
public class PdfWatermarkTool
{
    /// <summary>
    ///     Handler registry for watermark operations.
    /// </summary>
    private readonly HandlerRegistry<Document> _handlerRegistry;

    /// <summary>
    ///     The session identity accessor for session isolation.
    /// </summary>
    private readonly ISessionIdentityAccessor? _identityAccessor;

    /// <summary>
    ///     The document session manager for managing in-memory document sessions.
    /// </summary>
    private readonly DocumentSessionManager? _sessionManager;

    /// <summary>
    ///     Initializes a new instance of the <see cref="PdfWatermarkTool" /> class.
    /// </summary>
    /// <param name="sessionManager">Optional session manager for in-memory document editing.</param>
    /// <param name="identityAccessor">Optional session identity accessor for session isolation.</param>
    public PdfWatermarkTool(DocumentSessionManager? sessionManager = null,
        ISessionIdentityAccessor? identityAccessor = null)
    {
        _sessionManager = sessionManager;
        _identityAccessor = identityAccessor;
        _handlerRegistry = HandlerRegistry<Document>.CreateFromNamespace("AsposeMcpServer.Handlers.Pdf.Watermark");
    }

    /// <summary>
    ///     Executes a PDF watermark operation (add).
    /// </summary>
    /// <param name="operation">The operation to perform: add.</param>
    /// <param name="path">PDF file path (required if no sessionId).</param>
    /// <param name="sessionId">Session ID for in-memory editing.</param>
    /// <param name="outputPath">Output file path (file mode only).</param>
    /// <param name="text">Watermark text (required for add).</param>
    /// <param name="opacity">Opacity (0.0 to 1.0).</param>
    /// <param name="fontSize">Font size in points.</param>
    /// <param name="fontName">Font name.</param>
    /// <param name="rotation">Rotation angle in degrees.</param>
    /// <param name="color">Watermark color name or hex code.</param>
    /// <param name="pageRange">Page range to apply watermark (e.g., '1,3,5-10').</param>
    /// <param name="isBackground">If true, watermark is placed behind text content.</param>
    /// <param name="horizontalAlignment">Horizontal alignment.</param>
    /// <param name="verticalAlignment">Vertical alignment.</param>
    /// <returns>A message indicating the result of the operation.</returns>
    /// <exception cref="ArgumentException">Thrown when required parameters are missing or the operation is unknown.</exception>
    [McpServerTool(
        Name = "pdf_watermark",
        Title = "PDF Watermark Operations",
        Destructive = false,
        Idempotent = false,
        OpenWorld = false,
        ReadOnly = false,
        UseStructuredContent = true)]
    [Description(@"Manage watermarks in PDF documents. Supports 1 operation: add.

Usage examples:
- Add watermark: pdf_watermark(operation='add', path='doc.pdf', text='CONFIDENTIAL', fontSize=72, opacity=0.3)
- Add colored watermark: pdf_watermark(operation='add', path='doc.pdf', text='URGENT', color='Red')
- Add watermark to specific pages: pdf_watermark(operation='add', path='doc.pdf', text='DRAFT', pageRange='1,3,5-10')
- Add background watermark: pdf_watermark(operation='add', path='doc.pdf', text='SAMPLE', isBackground=true)")]
    public object Execute(
        [Description("Operation: add")] string operation,
        [Description("PDF file path (required if no sessionId)")]
        string? path = null,
        [Description("Session ID for in-memory editing")]
        string? sessionId = null,
        [Description("Output file path (file mode only)")]
        string? outputPath = null,
        [Description("Watermark text (required for add)")]
        string? text = null,
        [Description("Opacity (0.0 to 1.0, default: 0.3)")]
        double opacity = 0.3,
        [Description("Font size in points (default: 72)")]
        double fontSize = 72,
        [Description("Font name (default: 'Arial')")]
        string fontName = "Arial",
        [Description("Rotation angle in degrees (default: 45)")]
        double rotation = 45,
        [Description(
            "Watermark color name (e.g., 'Red', 'Blue', 'Gray') or hex code (e.g., '#FF0000'). Default: 'Gray'")]
        string color = "Gray",
        [Description("Page range to apply watermark (e.g., '1,3,5-10'). If not specified, applies to all pages")]
        string? pageRange = null,
        [Description("If true, watermark is placed behind text content. Default: false")]
        bool isBackground = false,
        [Description("Horizontal alignment (default: Center)")]
        string horizontalAlignment = "Center",
        [Description("Vertical alignment (default: Center)")]
        string verticalAlignment = "Center")
    {
        using var ctx = DocumentContext<Document>.Create(_sessionManager, sessionId, path, _identityAccessor);

        var parameters = BuildParameters(text, opacity, fontSize, fontName, rotation, color,
            pageRange, isBackground, horizontalAlignment, verticalAlignment);

        var handler = _handlerRegistry.GetHandler(operation);

        var operationContext = new OperationContext<Document>
        {
            Document = ctx.Document,
            SessionManager = _sessionManager,
            IdentityAccessor = _identityAccessor,
            SessionId = sessionId,
            SourcePath = path,
            OutputPath = outputPath
        };

        var result = handler.Execute(operationContext, parameters);

        if (operationContext.IsModified)
            ctx.Save(outputPath);

        return ResultHelper.FinalizeResult((dynamic)result, ctx, outputPath);
    }

    /// <summary>
    ///     Builds OperationParameters from method parameters.
    /// </summary>
    /// <param name="text">The watermark text.</param>
    /// <param name="opacity">The opacity (0.0 to 1.0).</param>
    /// <param name="fontSize">The font size in points.</param>
    /// <param name="fontName">The font name.</param>
    /// <param name="rotation">The rotation angle in degrees.</param>
    /// <param name="color">The watermark color name or hex code.</param>
    /// <param name="pageRange">The page range to apply watermark.</param>
    /// <param name="isBackground">Whether watermark is placed behind text content.</param>
    /// <param name="horizontalAlignment">The horizontal alignment.</param>
    /// <param name="verticalAlignment">The vertical alignment.</param>
    /// <returns>OperationParameters configured for the watermark operation.</returns>
    private static OperationParameters BuildParameters(
        string? text,
        double opacity,
        double fontSize,
        string fontName,
        double rotation,
        string color,
        string? pageRange,
        bool isBackground,
        string horizontalAlignment,
        string verticalAlignment)
    {
        var parameters = new OperationParameters();

        if (text != null) parameters.Set("text", text);
        parameters.Set("opacity", opacity);
        parameters.Set("fontSize", fontSize);
        parameters.Set("fontName", fontName);
        parameters.Set("rotation", rotation);
        parameters.Set("color", color);
        if (pageRange != null) parameters.Set("pageRange", pageRange);
        parameters.Set("isBackground", isBackground);
        parameters.Set("horizontalAlignment", horizontalAlignment);
        parameters.Set("verticalAlignment", verticalAlignment);

        return parameters;
    }
}
