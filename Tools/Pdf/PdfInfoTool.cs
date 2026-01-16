using System.ComponentModel;
using Aspose.Pdf;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Session;
using ModelContextProtocol.Server;

namespace AsposeMcpServer.Tools.Pdf;

/// <summary>
///     Tool for getting content and statistics from PDF documents
/// </summary>
[McpServerToolType]
public class PdfInfoTool
{
    /// <summary>
    ///     Handler registry for info operations.
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
    ///     Initializes a new instance of the <see cref="PdfInfoTool" /> class.
    /// </summary>
    /// <param name="sessionManager">Optional session manager for in-memory document editing.</param>
    /// <param name="identityAccessor">Optional session identity accessor for session isolation.</param>
    public PdfInfoTool(DocumentSessionManager? sessionManager = null, ISessionIdentityAccessor? identityAccessor = null)
    {
        _sessionManager = sessionManager;
        _identityAccessor = identityAccessor;
        _handlerRegistry = HandlerRegistry<Document>.CreateFromNamespace("AsposeMcpServer.Handlers.Pdf.Info");
    }

    /// <summary>
    ///     Executes a PDF info operation (get_content, get_statistics).
    /// </summary>
    /// <param name="operation">The operation to perform: get_content, get_statistics.</param>
    /// <param name="path">PDF file path (required if no sessionId).</param>
    /// <param name="sessionId">Session ID for in-memory editing.</param>
    /// <param name="pageIndex">Page index (1-based, optional for get_content).</param>
    /// <param name="maxPages">Maximum pages to extract (for get_content without pageIndex, default: 100).</param>
    /// <returns>A JSON string containing content or statistics data.</returns>
    /// <exception cref="ArgumentException">Thrown when required parameters are missing or the operation is unknown.</exception>
    [McpServerTool(Name = "pdf_info")]
    [Description(@"Get content and statistics from PDF documents. Supports 2 operations: get_content, get_statistics.

Usage examples:
- Get content from page: pdf_info(operation='get_content', path='doc.pdf', pageIndex=1)
- Get content from all pages: pdf_info(operation='get_content', path='doc.pdf')
- Get content with limit: pdf_info(operation='get_content', path='doc.pdf', maxPages=50)
- Get statistics: pdf_info(operation='get_statistics', path='doc.pdf')")]
    public string Execute(
        [Description(@"Operation to perform.
- 'get_content': Get text content from page(s) (required params: path)
- 'get_statistics': Get document statistics (required params: path)")]
        string operation,
        [Description("PDF file path (required if no sessionId)")]
        string? path = null,
        [Description("Session ID for in-memory editing")]
        string? sessionId = null,
        [Description("Page index (1-based, optional for get_content, extracts all if not specified)")]
        int? pageIndex = null,
        [Description("Maximum pages to extract (for get_content without pageIndex, default: 100)")]
        int maxPages = 100)
    {
        using var ctx = DocumentContext<Document>.Create(_sessionManager, sessionId, path, _identityAccessor);

        var parameters = BuildParameters(pageIndex, maxPages);

        var handler = _handlerRegistry.GetHandler(operation);

        var operationContext = new OperationContext<Document>
        {
            Document = ctx.Document,
            SessionManager = _sessionManager,
            IdentityAccessor = _identityAccessor,
            SessionId = sessionId,
            SourcePath = path
        };

        return handler.Execute(operationContext, parameters);
    }

    /// <summary>
    ///     Builds OperationParameters from method parameters.
    /// </summary>
    /// <param name="pageIndex">The page index (1-based).</param>
    /// <param name="maxPages">The maximum pages to extract.</param>
    /// <returns>OperationParameters configured for the info operation.</returns>
    private static OperationParameters BuildParameters(int? pageIndex, int maxPages)
    {
        var parameters = new OperationParameters();

        if (pageIndex.HasValue)
            parameters.Set("pageIndex", pageIndex.Value);

        parameters.Set("maxPages", maxPages);

        return parameters;
    }
}
