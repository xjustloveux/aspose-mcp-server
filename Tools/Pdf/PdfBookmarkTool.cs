using System.ComponentModel;
using Aspose.Pdf;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Session;
using AsposeMcpServer.Helpers;
using ModelContextProtocol.Server;

namespace AsposeMcpServer.Tools.Pdf;

/// <summary>
///     Tool for managing bookmarks in PDF documents (add, delete, edit, get)
/// </summary>
[ToolHandlerMapping("AsposeMcpServer.Handlers.Pdf.Bookmark")]
[McpServerToolType]
public class PdfBookmarkTool
{
    /// <summary>
    ///     Handler registry for bookmark operations.
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
    ///     Initializes a new instance of the <see cref="PdfBookmarkTool" /> class.
    /// </summary>
    /// <param name="sessionManager">Optional session manager for in-memory document editing.</param>
    /// <param name="identityAccessor">Optional session identity accessor for session isolation.</param>
    public PdfBookmarkTool(DocumentSessionManager? sessionManager = null,
        ISessionIdentityAccessor? identityAccessor = null)
    {
        _sessionManager = sessionManager;
        _identityAccessor = identityAccessor;
        _handlerRegistry = HandlerRegistry<Document>.CreateFromNamespace("AsposeMcpServer.Handlers.Pdf.Bookmark");
    }

    /// <summary>
    ///     Executes a PDF bookmark operation (add, delete, edit, get).
    /// </summary>
    /// <param name="operation">The operation to perform: add, delete, edit, get.</param>
    /// <param name="path">PDF file path (required if no sessionId).</param>
    /// <param name="sessionId">Session ID for in-memory editing.</param>
    /// <param name="outputPath">Output file path (file mode only).</param>
    /// <param name="title">Bookmark title (required for add, edit).</param>
    /// <param name="pageIndex">Target page index (1-based, required for add, edit).</param>
    /// <param name="bookmarkIndex">Bookmark index (1-based, required for delete, edit).</param>
    /// <returns>A message indicating the result of the operation, or JSON data for get operations.</returns>
    /// <exception cref="ArgumentException">Thrown when required parameters are missing or the operation is unknown.</exception>
    [McpServerTool(
        Name = "pdf_bookmark",
        Title = "PDF Bookmark Operations",
        Destructive = true,
        Idempotent = false,
        OpenWorld = false,
        ReadOnly = false,
        UseStructuredContent = true)]
    [Description(@"Manage bookmarks in PDF documents. Supports 4 operations: add, delete, edit, get.

Usage examples:
- Add bookmark: pdf_bookmark(operation='add', path='doc.pdf', title='Chapter 1', pageIndex=1)
- Delete bookmark: pdf_bookmark(operation='delete', path='doc.pdf', bookmarkIndex=0)
- Edit bookmark: pdf_bookmark(operation='edit', path='doc.pdf', bookmarkIndex=0, title='Updated Title', pageIndex=2)
- Get bookmarks: pdf_bookmark(operation='get', path='doc.pdf')")]
    public object Execute(
        [Description(@"Operation to perform.
- 'add': Add a bookmark (required params: path, title, pageIndex)
- 'delete': Delete a bookmark (required params: path, bookmarkIndex)
- 'edit': Edit a bookmark (required params: path, bookmarkIndex, title, pageIndex)
- 'get': Get all bookmarks (required params: path)")]
        string operation,
        [Description("PDF file path (required if no sessionId)")]
        string? path = null,
        [Description("Session ID for in-memory editing")]
        string? sessionId = null,
        [Description("Output file path (file mode only)")]
        string? outputPath = null,
        [Description("Bookmark title (required for add, edit)")]
        string? title = null,
        [Description("Target page index (1-based, required for add, edit)")]
        int? pageIndex = null,
        [Description("Bookmark index (1-based, required for delete, edit, optional for get)")]
        int? bookmarkIndex = null)
    {
        using var ctx = DocumentContext<Document>.Create(_sessionManager, sessionId, path, _identityAccessor);

        var parameters = BuildParameters(operation, title, pageIndex, bookmarkIndex);

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
        int? pageIndex,
        int? bookmarkIndex)
    {
        return operation.ToLowerInvariant() switch
        {
            "add" => BuildAddParameters(title, pageIndex),
            "delete" => BuildDeleteParameters(bookmarkIndex),
            "edit" => BuildEditParameters(bookmarkIndex, title, pageIndex),
            _ => new OperationParameters()
        };
    }

    /// <summary>
    ///     Builds parameters for the add bookmark operation.
    /// </summary>
    /// <param name="title">The bookmark title.</param>
    /// <param name="pageIndex">The target page index (1-based).</param>
    /// <returns>OperationParameters configured for adding a bookmark.</returns>
    private static OperationParameters BuildAddParameters(string? title, int? pageIndex)
    {
        var parameters = new OperationParameters();
        if (title != null) parameters.Set("title", title);
        if (pageIndex.HasValue) parameters.Set("pageIndex", pageIndex.Value);
        return parameters;
    }

    /// <summary>
    ///     Builds parameters for the delete bookmark operation.
    /// </summary>
    /// <param name="bookmarkIndex">The bookmark index (1-based) to delete.</param>
    /// <returns>OperationParameters configured for deleting a bookmark.</returns>
    private static OperationParameters BuildDeleteParameters(int? bookmarkIndex)
    {
        var parameters = new OperationParameters();
        if (bookmarkIndex.HasValue) parameters.Set("bookmarkIndex", bookmarkIndex.Value);
        return parameters;
    }

    /// <summary>
    ///     Builds parameters for the edit bookmark operation.
    /// </summary>
    /// <param name="bookmarkIndex">The bookmark index (1-based) to edit.</param>
    /// <param name="title">The new bookmark title.</param>
    /// <param name="pageIndex">The new target page index (1-based).</param>
    /// <returns>OperationParameters configured for editing a bookmark.</returns>
    private static OperationParameters BuildEditParameters(int? bookmarkIndex, string? title, int? pageIndex)
    {
        var parameters = new OperationParameters();
        if (bookmarkIndex.HasValue) parameters.Set("bookmarkIndex", bookmarkIndex.Value);
        if (title != null) parameters.Set("title", title);
        if (pageIndex.HasValue) parameters.Set("pageIndex", pageIndex.Value);
        return parameters;
    }
}
