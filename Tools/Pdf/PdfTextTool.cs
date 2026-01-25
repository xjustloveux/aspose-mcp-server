using System.ComponentModel;
using Aspose.Pdf;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Session;
using AsposeMcpServer.Helpers;
using ModelContextProtocol.Server;

namespace AsposeMcpServer.Tools.Pdf;

/// <summary>
///     Tool for managing text in PDF documents (add, edit, extract)
/// </summary>
[ToolHandlerMapping("AsposeMcpServer.Handlers.Pdf.Text")]
[McpServerToolType]
public class PdfTextTool
{
    /// <summary>
    ///     Handler registry for text operations.
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
    ///     Initializes a new instance of the <see cref="PdfTextTool" /> class.
    /// </summary>
    /// <param name="sessionManager">Optional session manager for in-memory document editing.</param>
    /// <param name="identityAccessor">Optional session identity accessor for session isolation.</param>
    public PdfTextTool(DocumentSessionManager? sessionManager = null, ISessionIdentityAccessor? identityAccessor = null)
    {
        _sessionManager = sessionManager;
        _identityAccessor = identityAccessor;
        _handlerRegistry = HandlerRegistry<Document>.CreateFromNamespace("AsposeMcpServer.Handlers.Pdf.Text");
    }

    /// <summary>
    ///     Executes a PDF text operation (add, edit, extract).
    /// </summary>
    /// <param name="operation">The operation to perform: add, edit, extract.</param>
    /// <param name="path">PDF file path (required if no sessionId).</param>
    /// <param name="sessionId">Session ID for in-memory editing.</param>
    /// <param name="outputPath">Output file path (file mode only).</param>
    /// <param name="pageIndex">Page index (1-based).</param>
    /// <param name="text">Text to add (required for add).</param>
    /// <param name="x">X position in PDF coordinates (for add).</param>
    /// <param name="y">Y position in PDF coordinates (for add).</param>
    /// <param name="fontName">Font name (for add).</param>
    /// <param name="fontSize">Font size (for add).</param>
    /// <param name="oldText">Text to replace (required for edit).</param>
    /// <param name="newText">New text (required for edit).</param>
    /// <param name="replaceAll">Replace all occurrences (for edit).</param>
    /// <param name="includeFontInfo">Include font information (for extract).</param>
    /// <param name="extractionMode">Text extraction mode (for extract).</param>
    /// <returns>A message indicating the result of the operation, or JSON data for extract operations.</returns>
    /// <exception cref="ArgumentException">Thrown when required parameters are missing or the operation is unknown.</exception>
    [McpServerTool(
        Name = "pdf_text",
        Title = "PDF Text Operations",
        Destructive = false,
        Idempotent = false,
        OpenWorld = false,
        ReadOnly = false,
        UseStructuredContent = true)]
    [Description(@"Manage text in PDF documents. Supports 3 operations: add, edit, extract.

Usage examples:
- Add text: pdf_text(operation='add', path='doc.pdf', pageIndex=1, text='Hello World', x=100, y=700)
- Add text with font: pdf_text(operation='add', path='doc.pdf', pageIndex=1, text='Hello', x=100, y=700, fontName='Arial', fontSize=14)
- Edit text: pdf_text(operation='edit', path='doc.pdf', pageIndex=1, oldText='old', newText='new')
- Edit all occurrences: pdf_text(operation='edit', path='doc.pdf', pageIndex=1, oldText='old', newText='new', replaceAll=true)
- Extract text: pdf_text(operation='extract', path='doc.pdf', pageIndex=1)
- Extract with font info: pdf_text(operation='extract', path='doc.pdf', pageIndex=1, includeFontInfo=true)
- Extract raw text: pdf_text(operation='extract', path='doc.pdf', pageIndex=1, extractionMode='raw')")]
    public object Execute(
        [Description("Operation: add, edit, extract")]
        string operation,
        [Description("PDF file path (required if no sessionId)")]
        string? path = null,
        [Description("Session ID for in-memory editing")]
        string? sessionId = null,
        [Description("Output file path (file mode only)")]
        string? outputPath = null,
        [Description("Page index (1-based)")] int pageIndex = 1,
        [Description("Text to add (required for add)")]
        string? text = null,
        [Description("X position in PDF coordinates (for add, default: 100)")]
        double x = 100,
        [Description("Y position in PDF coordinates (for add, default: 700)")]
        double y = 700,
        [Description("Font name (for add, default: 'Arial')")]
        string fontName = "Arial",
        [Description("Font size (for add, default: 12)")]
        double fontSize = 12,
        [Description("Text to replace (required for edit)")]
        string? oldText = null,
        [Description("New text (required for edit)")]
        string? newText = null,
        [Description("Replace all occurrences (for edit, default: false)")]
        bool replaceAll = false,
        [Description("Include font information (for extract, default: false)")]
        bool includeFontInfo = false,
        [Description("Text extraction mode (for extract, default: 'pure')")]
        string extractionMode = "pure")
    {
        using var ctx = DocumentContext<Document>.Create(_sessionManager, sessionId, path, _identityAccessor);

        var parameters = BuildParameters(operation, pageIndex, text, x, y, fontName, fontSize,
            oldText, newText, replaceAll, includeFontInfo, extractionMode);

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

        if (string.Equals(operation, "extract", StringComparison.OrdinalIgnoreCase))
            return ResultHelper.FinalizeResult((dynamic)result, ctx, outputPath);

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
        int pageIndex,
        string? text,
        double x,
        double y,
        string fontName,
        double fontSize,
        string? oldText,
        string? newText,
        bool replaceAll,
        bool includeFontInfo,
        string extractionMode)
    {
        return operation.ToLowerInvariant() switch
        {
            "add" => BuildAddParameters(pageIndex, text, x, y, fontName, fontSize),
            "edit" => BuildEditParameters(pageIndex, oldText, newText, replaceAll),
            "extract" => BuildExtractParameters(pageIndex, includeFontInfo, extractionMode),
            _ => new OperationParameters()
        };
    }

    /// <summary>
    ///     Builds parameters for the add text operation.
    /// </summary>
    /// <param name="pageIndex">The page index (1-based).</param>
    /// <param name="text">The text to add.</param>
    /// <param name="x">X position in PDF coordinates.</param>
    /// <param name="y">Y position in PDF coordinates.</param>
    /// <param name="fontName">The font name.</param>
    /// <param name="fontSize">The font size.</param>
    /// <returns>OperationParameters configured for adding text.</returns>
    private static OperationParameters BuildAddParameters(int pageIndex, string? text, double x, double y,
        string fontName, double fontSize)
    {
        var parameters = new OperationParameters();
        parameters.Set("pageIndex", pageIndex);
        if (text != null) parameters.Set("text", text);
        parameters.Set("x", x);
        parameters.Set("y", y);
        parameters.Set("fontName", fontName);
        parameters.Set("fontSize", fontSize);
        return parameters;
    }

    /// <summary>
    ///     Builds parameters for the edit (replace) text operation.
    /// </summary>
    /// <param name="pageIndex">The page index (1-based).</param>
    /// <param name="oldText">The text to replace.</param>
    /// <param name="newText">The new text.</param>
    /// <param name="replaceAll">Whether to replace all occurrences.</param>
    /// <returns>OperationParameters configured for editing text.</returns>
    private static OperationParameters BuildEditParameters(int pageIndex, string? oldText, string? newText,
        bool replaceAll)
    {
        var parameters = new OperationParameters();
        parameters.Set("pageIndex", pageIndex);
        if (oldText != null) parameters.Set("oldText", oldText);
        if (newText != null) parameters.Set("newText", newText);
        parameters.Set("replaceAll", replaceAll);
        return parameters;
    }

    /// <summary>
    ///     Builds parameters for the extract text operation.
    /// </summary>
    /// <param name="pageIndex">The page index (1-based).</param>
    /// <param name="includeFontInfo">Whether to include font information.</param>
    /// <param name="extractionMode">The text extraction mode.</param>
    /// <returns>OperationParameters configured for extracting text.</returns>
    private static OperationParameters BuildExtractParameters(int pageIndex, bool includeFontInfo,
        string extractionMode)
    {
        var parameters = new OperationParameters();
        parameters.Set("pageIndex", pageIndex);
        parameters.Set("includeFontInfo", includeFontInfo);
        parameters.Set("extractionMode", extractionMode);
        return parameters;
    }
}
