using System.ComponentModel;
using Aspose.Pdf;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Session;
using AsposeMcpServer.Handlers.Pdf.Page;
using ModelContextProtocol.Server;

namespace AsposeMcpServer.Tools.Pdf;

/// <summary>
///     Tool for managing pages in PDF documents (add, delete, insert, extract, rotate, resize)
/// </summary>
[McpServerToolType]
public class PdfPageTool
{
    /// <summary>
    ///     Handler registry for page operations.
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
    ///     Initializes a new instance of the <see cref="PdfPageTool" /> class.
    /// </summary>
    /// <param name="sessionManager">Optional session manager for in-memory document editing.</param>
    /// <param name="identityAccessor">Optional session identity accessor for session isolation.</param>
    public PdfPageTool(DocumentSessionManager? sessionManager = null, ISessionIdentityAccessor? identityAccessor = null)
    {
        _sessionManager = sessionManager;
        _identityAccessor = identityAccessor;
        _handlerRegistry = PdfPageHandlerRegistry.Create();
    }

    /// <summary>
    ///     Executes a PDF page operation (add, delete, rotate, get_details, get_info).
    /// </summary>
    /// <param name="operation">The operation to perform: add, delete, rotate, get_details, get_info.</param>
    /// <param name="path">PDF file path (required if no sessionId).</param>
    /// <param name="sessionId">Session ID for in-memory editing.</param>
    /// <param name="outputPath">Output file path (optional, defaults to overwrite input).</param>
    /// <param name="count">Number of pages to add (for add, default: 1).</param>
    /// <param name="insertAt">Position to insert pages (1-based, for add, optional).</param>
    /// <param name="width">Page width in points (for add, optional).</param>
    /// <param name="height">Page height in points (for add, optional).</param>
    /// <param name="pageIndex">Page index (1-based, required for delete, rotate, get_details).</param>
    /// <param name="rotation">Rotation angle in degrees: 0, 90, 180, 270 (for rotate, required).</param>
    /// <param name="pageIndices">Array of page indices to rotate (1-based, for rotate, optional).</param>
    /// <returns>A message indicating the result of the operation, or JSON data for get operations.</returns>
    /// <exception cref="ArgumentException">Thrown when required parameters are missing or the operation is unknown.</exception>
    [McpServerTool(Name = "pdf_page")]
    [Description(@"Manage pages in PDF documents. Supports 5 operations: add, delete, rotate, get_details, get_info.

Usage examples:
- Add page: pdf_page(operation='add', path='doc.pdf', count=1)
- Delete page: pdf_page(operation='delete', path='doc.pdf', pageIndex=1)
- Rotate page: pdf_page(operation='rotate', path='doc.pdf', pageIndex=1, rotation=90)
- Get page details: pdf_page(operation='get_details', path='doc.pdf', pageIndex=1)
- Get page info: pdf_page(operation='get_info', path='doc.pdf')")]
    public string Execute(
        [Description(@"Operation to perform.
- 'add': Add page(s) (required params: path)
- 'delete': Delete a page (required params: path, pageIndex)
- 'rotate': Rotate a page (required params: path, pageIndex, rotation)
- 'get_details': Get page details (required params: path, pageIndex)
- 'get_info': Get all pages info (required params: path)")]
        string operation,
        [Description("PDF file path (required if no sessionId)")]
        string? path = null,
        [Description("Session ID for in-memory editing")]
        string? sessionId = null,
        [Description("Output file path (optional, defaults to overwrite input)")]
        string? outputPath = null,
        [Description("Number of pages to add (for add, default: 1)")]
        int count = 1,
        [Description("Position to insert pages (1-based, for add, optional, default: append at end)")]
        int? insertAt = null,
        [Description("Page width in points (for add, optional)")]
        double? width = null,
        [Description("Page height in points (for add, optional)")]
        double? height = null,
        [Description("Page index (1-based, required for delete, rotate, get_details)")]
        int pageIndex = 0,
        [Description("Rotation angle in degrees: 0, 90, 180, 270 (for rotate, required)")]
        int rotation = 0,
        [Description("Array of page indices to rotate (1-based, for rotate, optional)")]
        int[]? pageIndices = null)
    {
        using var ctx = DocumentContext<Document>.Create(_sessionManager, sessionId, path, _identityAccessor);

        var parameters = BuildParameters(operation, count, insertAt, width, height, pageIndex, rotation, pageIndices);

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

        if (operation.ToLowerInvariant() is "get_details" or "get_info")
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
        int count,
        int? insertAt,
        double? width,
        double? height,
        int pageIndex,
        int rotation,
        int[]? pageIndices)
    {
        var parameters = new OperationParameters();

        switch (operation.ToLowerInvariant())
        {
            case "add":
                parameters.Set("count", count);
                if (insertAt.HasValue) parameters.Set("insertAt", insertAt.Value);
                if (width.HasValue) parameters.Set("width", width.Value);
                if (height.HasValue) parameters.Set("height", height.Value);
                break;

            case "delete":
                parameters.Set("pageIndex", pageIndex);
                break;

            case "rotate":
                parameters.Set("pageIndex", pageIndex);
                parameters.Set("rotation", rotation);
                if (pageIndices != null) parameters.Set("pageIndices", pageIndices);
                break;

            case "get_details":
                parameters.Set("pageIndex", pageIndex);
                break;

            case "get_info":
                break;
        }

        return parameters;
    }
}
