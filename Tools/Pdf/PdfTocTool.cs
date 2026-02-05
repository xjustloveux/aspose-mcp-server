using System.ComponentModel;
using Aspose.Pdf;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Session;
using AsposeMcpServer.Helpers;
using ModelContextProtocol.Server;

namespace AsposeMcpServer.Tools.Pdf;

/// <summary>
///     Tool for managing table of contents in PDF documents (generate, get, remove).
/// </summary>
[ToolHandlerMapping("AsposeMcpServer.Handlers.Pdf.Toc")]
[McpServerToolType]
public class PdfTocTool
{
    /// <summary>
    ///     Handler registry for TOC operations.
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
    ///     Initializes a new instance of the <see cref="PdfTocTool" /> class.
    /// </summary>
    /// <param name="sessionManager">Optional session manager for in-memory document editing.</param>
    /// <param name="identityAccessor">Optional session identity accessor for session isolation.</param>
    public PdfTocTool(DocumentSessionManager? sessionManager = null,
        ISessionIdentityAccessor? identityAccessor = null)
    {
        _sessionManager = sessionManager;
        _identityAccessor = identityAccessor;
        _handlerRegistry = HandlerRegistry<Document>.CreateFromNamespace("AsposeMcpServer.Handlers.Pdf.Toc");
    }

    /// <summary>
    ///     Executes a PDF table of contents operation (generate, get, remove).
    /// </summary>
    /// <param name="operation">The operation to perform: generate, get, remove.</param>
    /// <param name="path">PDF file path (required if no sessionId).</param>
    /// <param name="sessionId">Session ID for in-memory editing.</param>
    /// <param name="outputPath">Output file path (file mode only).</param>
    /// <param name="title">TOC title (used for generate operation).</param>
    /// <param name="depth">Maximum heading depth to include (used for generate operation).</param>
    /// <param name="tocPage">Page position to insert the TOC (1-based, used for generate operation).</param>
    /// <returns>A message indicating the result of the operation, or JSON data for get operations.</returns>
    /// <exception cref="ArgumentException">Thrown when required parameters are missing or the operation is unknown.</exception>
    [McpServerTool(
        Name = "pdf_toc",
        Title = "PDF Table of Contents Operations",
        Destructive = true,
        Idempotent = false,
        OpenWorld = false,
        ReadOnly = false,
        UseStructuredContent = true)]
    [Description(
        @"Generate, read, and remove table of contents in PDF documents. Supports 3 operations: generate, get, remove.

Usage examples:
- Generate TOC: pdf_toc(operation='generate', path='doc.pdf', outputPath='out.pdf')
- Generate with options: pdf_toc(operation='generate', path='doc.pdf', outputPath='out.pdf', title='Contents', depth=2, tocPage=1)
- Get TOC entries: pdf_toc(operation='get', path='doc.pdf')
- Remove TOC: pdf_toc(operation='remove', path='doc.pdf', outputPath='out.pdf')")]
    public object Execute(
        [Description(@"Operation to perform.
- 'generate': Generate a table of contents page (optional params: title, depth, tocPage)
- 'get': Get TOC entries from outlines (required params: path)
- 'remove': Remove TOC pages (required params: path)")]
        string operation,
        [Description("PDF file path (required if no sessionId)")]
        string? path = null,
        [Description("Session ID for in-memory editing")]
        string? sessionId = null,
        [Description("Output file path (file mode only)")]
        string? outputPath = null,
        [Description("TOC title (default: 'Table of Contents')")]
        string? title = null,
        [Description("Maximum heading depth to include (default: 3)")]
        int? depth = null,
        [Description("Page position to insert TOC (1-based, default: 1)")]
        int? tocPage = null)
    {
        using var ctx = DocumentContext<Document>.Create(_sessionManager, sessionId, path, _identityAccessor);

        var parameters = BuildParameters(operation, title, depth, tocPage);

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

        if (string.Equals(operation, "get", StringComparison.OrdinalIgnoreCase))
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
        string? title,
        int? depth,
        int? tocPage)
    {
        return operation.ToLowerInvariant() switch
        {
            "generate" => BuildGenerateParameters(title, depth, tocPage),
            _ => new OperationParameters()
        };
    }

    /// <summary>
    ///     Builds parameters for the generate TOC operation.
    /// </summary>
    /// <param name="title">The TOC title.</param>
    /// <param name="depth">The maximum heading depth.</param>
    /// <param name="tocPage">The page position to insert the TOC.</param>
    /// <returns>OperationParameters configured for generating a TOC.</returns>
    private static OperationParameters BuildGenerateParameters(string? title, int? depth, int? tocPage)
    {
        var parameters = new OperationParameters();
        if (title != null) parameters.Set("title", title);
        if (depth.HasValue) parameters.Set("depth", depth.Value);
        if (tocPage.HasValue) parameters.Set("tocPage", tocPage.Value);
        return parameters;
    }
}
