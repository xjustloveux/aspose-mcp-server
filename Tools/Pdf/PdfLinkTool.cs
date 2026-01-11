using System.ComponentModel;
using Aspose.Pdf;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Session;
using AsposeMcpServer.Handlers.Pdf.Link;
using ModelContextProtocol.Server;

namespace AsposeMcpServer.Tools.Pdf;

/// <summary>
///     Tool for managing links in PDF documents (add, delete, edit, get)
/// </summary>
[McpServerToolType]
public class PdfLinkTool
{
    /// <summary>
    ///     Handler registry for link operations.
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
    ///     Initializes a new instance of the <see cref="PdfLinkTool" /> class.
    /// </summary>
    /// <param name="sessionManager">Optional session manager for in-memory document editing.</param>
    /// <param name="identityAccessor">Optional session identity accessor for session isolation.</param>
    public PdfLinkTool(DocumentSessionManager? sessionManager = null, ISessionIdentityAccessor? identityAccessor = null)
    {
        _sessionManager = sessionManager;
        _identityAccessor = identityAccessor;
        _handlerRegistry = PdfLinkHandlerRegistry.Create();
    }

    /// <summary>
    ///     Executes a PDF link operation (add, delete, edit, get).
    /// </summary>
    /// <param name="operation">The operation to perform: add, delete, edit, get.</param>
    /// <param name="path">PDF file path (required if no sessionId).</param>
    /// <param name="sessionId">Session ID for in-memory editing.</param>
    /// <param name="outputPath">Output file path (optional, defaults to overwrite input).</param>
    /// <param name="pageIndex">Page index (1-based, required for add, delete, edit).</param>
    /// <param name="linkIndex">Link index (0-based, required for delete, edit).</param>
    /// <param name="x">X position of link area in PDF coordinates (required for add).</param>
    /// <param name="y">Y position of link area in PDF coordinates (required for add).</param>
    /// <param name="width">Width of link area in PDF points (required for add).</param>
    /// <param name="height">Height of link area in PDF points (required for add).</param>
    /// <param name="url">URL to link to (for add, edit).</param>
    /// <param name="targetPage">Target page number (1-based, for add, edit).</param>
    /// <returns>A message indicating the result of the operation, or JSON data for get operations.</returns>
    /// <exception cref="ArgumentException">Thrown when required parameters are missing or the operation is unknown.</exception>
    [McpServerTool(Name = "pdf_link")]
    [Description(@"Manage links in PDF documents. Supports 4 operations: add, delete, edit, get.

Usage examples:
- Add link: pdf_link(operation='add', path='doc.pdf', pageIndex=1, x=100, y=100, width=200, height=30, url='https://example.com')
- Delete link: pdf_link(operation='delete', path='doc.pdf', pageIndex=1, linkIndex=0)
- Edit link: pdf_link(operation='edit', path='doc.pdf', pageIndex=1, linkIndex=0, url='https://newurl.com')
- Get links: pdf_link(operation='get', path='doc.pdf', pageIndex=1)")]
    public string Execute(
        [Description(@"Operation to perform.
- 'add': Add a link (required params: path, pageIndex, x, y, width, height, url)
- 'delete': Delete a link (required params: path, pageIndex, linkIndex)
- 'edit': Edit a link (required params: path, pageIndex, linkIndex, url)
- 'get': Get all links (required params: path, pageIndex)")]
        string operation,
        [Description("PDF file path (required if no sessionId)")]
        string? path = null,
        [Description("Session ID for in-memory editing")]
        string? sessionId = null,
        [Description("Output file path (optional, defaults to overwrite input)")]
        string? outputPath = null,
        [Description("Page index (1-based, required for add, delete, edit)")]
        int pageIndex = 0,
        [Description("Link index (0-based, required for delete, edit)")]
        int linkIndex = 0,
        [Description("X position of link area in PDF coordinates, origin at bottom-left corner (required for add)")]
        double x = 0,
        [Description("Y position of link area in PDF coordinates, origin at bottom-left corner (required for add)")]
        double y = 0,
        [Description("Width of link area in PDF points (required for add)")]
        double width = 0,
        [Description("Height of link area in PDF points (required for add)")]
        double height = 0,
        [Description("URL to link to (for add, edit)")]
        string? url = null,
        [Description("Target page number (1-based, for add, edit)")]
        int? targetPage = null)
    {
        using var ctx = DocumentContext<Document>.Create(_sessionManager, sessionId, path, _identityAccessor);

        var parameters = BuildParameters(operation, pageIndex, linkIndex, x, y, width, height, url, targetPage);

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

        if (operation.ToLowerInvariant() == "get")
            return result;

        if (operationContext.IsModified)
            ctx.Save(outputPath);

        return $"{result}\n{ctx.GetOutputMessage(outputPath)}";
    }

    /// <summary>
    ///     Builds OperationParameters from method parameters.
    /// </summary>
    private static OperationParameters BuildParameters(
        string operation,
        int pageIndex,
        int linkIndex,
        double x,
        double y,
        double width,
        double height,
        string? url,
        int? targetPage)
    {
        var parameters = new OperationParameters();

        switch (operation.ToLowerInvariant())
        {
            case "add":
                parameters.Set("pageIndex", pageIndex);
                parameters.Set("x", x);
                parameters.Set("y", y);
                parameters.Set("width", width);
                parameters.Set("height", height);
                if (url != null) parameters.Set("url", url);
                if (targetPage.HasValue) parameters.Set("targetPage", targetPage.Value);
                break;

            case "delete":
                parameters.Set("pageIndex", pageIndex);
                parameters.Set("linkIndex", linkIndex);
                break;

            case "edit":
                parameters.Set("pageIndex", pageIndex);
                parameters.Set("linkIndex", linkIndex);
                if (url != null) parameters.Set("url", url);
                if (targetPage.HasValue) parameters.Set("targetPage", targetPage.Value);
                break;

            case "get":
                if (pageIndex > 0) parameters.Set("pageIndex", pageIndex);
                break;
        }

        return parameters;
    }
}
