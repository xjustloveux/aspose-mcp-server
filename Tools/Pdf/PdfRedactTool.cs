using System.ComponentModel;
using Aspose.Pdf;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Session;
using AsposeMcpServer.Helpers;
using ModelContextProtocol.Server;

namespace AsposeMcpServer.Tools.Pdf;

/// <summary>
///     Tool for redacting (blacking out) text or areas on PDF pages
/// </summary>
[ToolHandlerMapping("AsposeMcpServer.Handlers.Pdf.Redact")]
[McpServerToolType]
public class PdfRedactTool
{
    /// <summary>
    ///     Handler registry for redact operations.
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
    ///     Initializes a new instance of the <see cref="PdfRedactTool" /> class.
    /// </summary>
    /// <param name="sessionManager">Optional session manager for in-memory document editing.</param>
    /// <param name="identityAccessor">Optional session identity accessor for session isolation.</param>
    public PdfRedactTool(DocumentSessionManager? sessionManager = null,
        ISessionIdentityAccessor? identityAccessor = null)
    {
        _sessionManager = sessionManager;
        _identityAccessor = identityAccessor;
        _handlerRegistry = HandlerRegistry<Document>.CreateFromNamespace("AsposeMcpServer.Handlers.Pdf.Redact");
    }

    /// <summary>
    ///     Executes a PDF redaction operation (area redaction or text search redaction).
    /// </summary>
    /// <param name="path">PDF file path (required if no sessionId).</param>
    /// <param name="sessionId">Session ID for in-memory editing.</param>
    /// <param name="outputPath">Output file path (file mode only).</param>
    /// <param name="pageIndex">Page index (1-based, optional for text search).</param>
    /// <param name="x">X position of redaction area in PDF coordinates (required for area redaction).</param>
    /// <param name="y">Y position of redaction area in PDF coordinates (required for area redaction).</param>
    /// <param name="width">Width of redaction area in PDF points (required for area redaction).</param>
    /// <param name="height">Height of redaction area in PDF points (required for area redaction).</param>
    /// <param name="textToRedact">Text to search and redact (required for text redaction).</param>
    /// <param name="caseSensitive">Whether text search is case sensitive (default: true).</param>
    /// <param name="fillColor">Fill color (optional, default: black).</param>
    /// <param name="overlayText">Text to display over the redacted area (optional).</param>
    /// <returns>A message indicating the result of the operation.</returns>
    /// <exception cref="ArgumentException">Thrown when required parameters are missing.</exception>
    [McpServerTool(
        Name = "pdf_redact",
        Title = "PDF Redaction Operations",
        Destructive = true,
        Idempotent = false,
        OpenWorld = false,
        ReadOnly = false,
        UseStructuredContent = true)]
    [Description(@"Redact (black out) text or area on PDF page. This permanently removes the underlying content.
Auto-detects mode: if textToRedact is provided, uses text search mode; otherwise uses area redaction mode.

Usage examples:
- Redact area: pdf_redact(path='doc.pdf', pageIndex=1, x=100, y=100, width=200, height=50)
- Redact with color: pdf_redact(path='doc.pdf', pageIndex=1, x=100, y=100, width=200, height=50, fillColor='255,0,0')
- Redact with overlay: pdf_redact(path='doc.pdf', pageIndex=1, x=100, y=100, width=200, height=50, overlayText='[REDACTED]')
- Redact by text search: pdf_redact(path='doc.pdf', textToRedact='confidential')
- Redact by text on page: pdf_redact(path='doc.pdf', pageIndex=1, textToRedact='secret', caseSensitive=false)")]
    public object Execute(
        [Description("PDF file path (required if no sessionId)")]
        string? path = null,
        [Description("Session ID for in-memory editing")]
        string? sessionId = null,
        [Description("Output file path (file mode only)")]
        string? outputPath = null,
        [Description("Page index (1-based, required for area redaction, optional for text search)")]
        int? pageIndex = null,
        [Description(
            "X position of redaction area in PDF coordinates, origin at bottom-left corner (required for area redaction)")]
        double? x = null,
        [Description(
            "Y position of redaction area in PDF coordinates, origin at bottom-left corner (required for area redaction)")]
        double? y = null,
        [Description("Width of redaction area in PDF points (required for area redaction)")]
        double? width = null,
        [Description("Height of redaction area in PDF points (required for area redaction)")]
        double? height = null,
        [Description("Text to search and redact (if provided, uses text search mode instead of area redaction)")]
        string? textToRedact = null,
        [Description("Whether text search is case sensitive (default: true, only for text mode)")]
        bool caseSensitive = true,
        [Description("Fill color (optional, default: black, format: 'R,G,B' or color name)")]
        string? fillColor = null,
        [Description("Text to display over the redacted area (optional, e.g., '[REDACTED]')")]
        string? overlayText = null)
    {
        using var ctx = DocumentContext<Document>.Create(_sessionManager, sessionId, path, _identityAccessor);

        // Auto-detect operation based on parameters
        var operation = !string.IsNullOrEmpty(textToRedact) ? "text" : "area";

        var parameters = BuildParameters(operation, pageIndex, x, y, width, height,
            textToRedact, caseSensitive, fillColor, overlayText);

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
    ///     Builds OperationParameters from method parameters using strategy pattern.
    ///     Parameters are documented on the Execute method.
    /// </summary>
    /// <returns>OperationParameters configured with all input values.</returns>
    private static OperationParameters BuildParameters(
        string operation,
        int? pageIndex,
        double? x,
        double? y,
        double? width,
        double? height,
        string? textToRedact,
        bool caseSensitive,
        string? fillColor,
        string? overlayText)
    {
        return operation.ToLowerInvariant() switch
        {
            "area" => BuildAreaParameters(pageIndex, x, y, width, height, fillColor, overlayText),
            "text" => BuildTextParameters(pageIndex, textToRedact, caseSensitive, fillColor, overlayText),
            _ => new OperationParameters()
        };
    }

    /// <summary>
    ///     Builds parameters for the area redaction operation.
    /// </summary>
    /// <param name="pageIndex">The page index (1-based).</param>
    /// <param name="x">X position of redaction area.</param>
    /// <param name="y">Y position of redaction area.</param>
    /// <param name="width">Width of redaction area.</param>
    /// <param name="height">Height of redaction area.</param>
    /// <param name="fillColor">Fill color for redaction.</param>
    /// <param name="overlayText">Text to display over redacted area.</param>
    /// <returns>OperationParameters configured for area redaction.</returns>
    private static OperationParameters BuildAreaParameters(int? pageIndex, double? x, double? y, double? width,
        double? height, string? fillColor, string? overlayText)
    {
        var parameters = new OperationParameters();
        if (pageIndex.HasValue) parameters.Set("pageIndex", pageIndex.Value);
        if (x.HasValue) parameters.Set("x", x.Value);
        if (y.HasValue) parameters.Set("y", y.Value);
        if (width.HasValue) parameters.Set("width", width.Value);
        if (height.HasValue) parameters.Set("height", height.Value);
        if (fillColor != null) parameters.Set("fillColor", fillColor);
        if (overlayText != null) parameters.Set("overlayText", overlayText);
        return parameters;
    }

    /// <summary>
    ///     Builds parameters for the text search redaction operation.
    /// </summary>
    /// <param name="pageIndex">The page index (1-based, optional).</param>
    /// <param name="textToRedact">The text to search and redact.</param>
    /// <param name="caseSensitive">Whether the search is case sensitive.</param>
    /// <param name="fillColor">Fill color for redaction.</param>
    /// <param name="overlayText">Text to display over redacted area.</param>
    /// <returns>OperationParameters configured for text redaction.</returns>
    private static OperationParameters BuildTextParameters(int? pageIndex, string? textToRedact, bool caseSensitive,
        string? fillColor, string? overlayText)
    {
        var parameters = new OperationParameters();
        if (pageIndex.HasValue) parameters.Set("pageIndex", pageIndex.Value);
        if (textToRedact != null) parameters.Set("textToRedact", textToRedact);
        parameters.Set("caseSensitive", caseSensitive);
        if (fillColor != null) parameters.Set("fillColor", fillColor);
        if (overlayText != null) parameters.Set("overlayText", overlayText);
        return parameters;
    }
}
