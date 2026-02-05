using System.ComponentModel;
using Aspose.Pdf;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Session;
using AsposeMcpServer.Helpers;
using ModelContextProtocol.Server;

namespace AsposeMcpServer.Tools.Pdf;

/// <summary>
///     Tool for managing pages in PDF documents (add, delete, insert, extract, rotate, resize)
/// </summary>
[ToolHandlerMapping("AsposeMcpServer.Handlers.Pdf.Page")]
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
        _handlerRegistry = HandlerRegistry<Document>.CreateFromNamespace("AsposeMcpServer.Handlers.Pdf.Page");
    }

    /// <summary>
    ///     Executes a PDF page operation (add, delete, rotate, crop, resize, get_details, get_info).
    /// </summary>
    /// <param name="operation">The operation to perform: add, delete, rotate, crop, resize, get_details, get_info.</param>
    /// <param name="path">PDF file path (required if no sessionId).</param>
    /// <param name="sessionId">Session ID for in-memory editing.</param>
    /// <param name="outputPath">Output file path (optional, defaults to overwrite input).</param>
    /// <param name="count">Number of pages to add (for add, default: 1).</param>
    /// <param name="insertAt">Position to insert pages (1-based, for add, optional).</param>
    /// <param name="width">Page width in points (for add, optional).</param>
    /// <param name="height">Page height in points (for add, optional).</param>
    /// <param name="x">X position in points (for crop, lower-left corner).</param>
    /// <param name="y">Y position in points (for crop, lower-left corner).</param>
    /// <param name="pageIndex">Page index (1-based, required for delete, rotate, get_details).</param>
    /// <param name="rotation">Rotation angle in degrees: 0, 90, 180, 270 (for rotate, required).</param>
    /// <param name="pageIndices">Array of page indices to rotate (1-based, for rotate, optional).</param>
    /// <returns>A message indicating the result of the operation, or JSON data for get operations.</returns>
    /// <exception cref="ArgumentException">Thrown when required parameters are missing or the operation is unknown.</exception>
    [McpServerTool(
        Name = "pdf_page",
        Title = "PDF Page Operations",
        Destructive = true,
        Idempotent = false,
        OpenWorld = false,
        ReadOnly = false,
        UseStructuredContent = true)]
    [Description(
        @"Manage pages in PDF documents. Supports 7 operations: add, delete, rotate, crop, resize, get_details, get_info.

Usage examples:
- Add page: pdf_page(operation='add', path='doc.pdf', count=1)
- Delete page: pdf_page(operation='delete', path='doc.pdf', pageIndex=1)
- Rotate page: pdf_page(operation='rotate', path='doc.pdf', pageIndex=1, rotation=90)
- Crop page: pdf_page(operation='crop', path='doc.pdf', pageIndex=1, x=50, y=50, width=400, height=600)
- Resize page: pdf_page(operation='resize', path='doc.pdf', pageIndex=1, width=595, height=842)
- Get page details: pdf_page(operation='get_details', path='doc.pdf', pageIndex=1)
- Get page info: pdf_page(operation='get_info', path='doc.pdf')")]
    public object Execute(
        [Description(@"Operation to perform.
- 'add': Add page(s) (required params: path)
- 'delete': Delete a page (required params: path, pageIndex)
- 'rotate': Rotate a page (required params: path, pageIndex, rotation)
- 'crop': Crop a page (required params: path, pageIndex, x, y, width, height)
- 'resize': Resize a page (required params: path, pageIndex, width, height)
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
        [Description("Page width in points (for add, crop, resize)")]
        double? width = null,
        [Description("Page height in points (for add, crop, resize)")]
        double? height = null,
        [Description("X position in points (for crop, lower-left corner)")]
        double? x = null,
        [Description("Y position in points (for crop, lower-left corner)")]
        double? y = null,
        [Description("Page index (1-based, required for delete, rotate, crop, resize, get_details)")]
        int pageIndex = 0,
        [Description("Rotation angle in degrees: 0, 90, 180, 270 (for rotate, required)")]
        int rotation = 0,
        [Description("Array of page indices to rotate (1-based, for rotate, optional)")]
        int[]? pageIndices = null)
    {
        using var ctx = DocumentContext<Document>.Create(_sessionManager, sessionId, path, _identityAccessor);

        var parameters = BuildParameters(operation, count, insertAt, width, height, x, y, pageIndex, rotation,
            pageIndices);

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
        int count,
        int? insertAt,
        double? width,
        double? height,
        double? x,
        double? y,
        int pageIndex,
        int rotation,
        int[]? pageIndices)
    {
        return operation.ToLowerInvariant() switch
        {
            "add" => BuildAddParameters(count, insertAt, width, height),
            "delete" => BuildDeleteParameters(pageIndex),
            "rotate" => BuildRotateParameters(pageIndex, rotation, pageIndices),
            "crop" => BuildCropParameters(pageIndex, x, y, width, height),
            "resize" => BuildResizeParameters(pageIndex, width, height),
            "get_details" => BuildGetDetailsParameters(pageIndex),
            _ => new OperationParameters()
        };
    }

    /// <summary>
    ///     Builds parameters for the add page operation.
    /// </summary>
    /// <param name="count">The number of pages to add.</param>
    /// <param name="insertAt">The position to insert pages (1-based).</param>
    /// <param name="width">The page width in points.</param>
    /// <param name="height">The page height in points.</param>
    /// <returns>OperationParameters configured for adding pages.</returns>
    private static OperationParameters BuildAddParameters(int count, int? insertAt, double? width, double? height)
    {
        var parameters = new OperationParameters();
        parameters.Set("count", count);
        if (insertAt.HasValue) parameters.Set("insertAt", insertAt.Value);
        if (width.HasValue) parameters.Set("width", width.Value);
        if (height.HasValue) parameters.Set("height", height.Value);
        return parameters;
    }

    /// <summary>
    ///     Builds parameters for the delete page operation.
    /// </summary>
    /// <param name="pageIndex">The page index (1-based) to delete.</param>
    /// <returns>OperationParameters configured for deleting a page.</returns>
    private static OperationParameters BuildDeleteParameters(int pageIndex)
    {
        var parameters = new OperationParameters();
        parameters.Set("pageIndex", pageIndex);
        return parameters;
    }

    /// <summary>
    ///     Builds parameters for the rotate page operation.
    /// </summary>
    /// <param name="pageIndex">The page index (1-based) to rotate.</param>
    /// <param name="rotation">The rotation angle in degrees (0, 90, 180, 270).</param>
    /// <param name="pageIndices">Array of page indices (1-based) to rotate.</param>
    /// <returns>OperationParameters configured for rotating pages.</returns>
    private static OperationParameters BuildRotateParameters(int pageIndex, int rotation, int[]? pageIndices)
    {
        var parameters = new OperationParameters();
        parameters.Set("pageIndex", pageIndex);
        parameters.Set("rotation", rotation);
        if (pageIndices != null) parameters.Set("pageIndices", pageIndices);
        return parameters;
    }

    /// <summary>
    ///     Builds parameters for the crop page operation.
    /// </summary>
    /// <param name="pageIndex">The page index (1-based) to crop.</param>
    /// <param name="x">The lower-left X coordinate.</param>
    /// <param name="y">The lower-left Y coordinate.</param>
    /// <param name="width">The crop width.</param>
    /// <param name="height">The crop height.</param>
    /// <returns>OperationParameters configured for cropping a page.</returns>
    private static OperationParameters BuildCropParameters(int pageIndex, double? x, double? y, double? width,
        double? height)
    {
        var parameters = new OperationParameters();
        parameters.Set("pageIndex", pageIndex);
        if (x.HasValue) parameters.Set("x", x.Value);
        if (y.HasValue) parameters.Set("y", y.Value);
        if (width.HasValue) parameters.Set("width", width.Value);
        if (height.HasValue) parameters.Set("height", height.Value);
        return parameters;
    }

    /// <summary>
    ///     Builds parameters for the resize page operation.
    /// </summary>
    /// <param name="pageIndex">The page index (1-based) to resize.</param>
    /// <param name="width">The new page width in points.</param>
    /// <param name="height">The new page height in points.</param>
    /// <returns>OperationParameters configured for resizing a page.</returns>
    private static OperationParameters BuildResizeParameters(int pageIndex, double? width, double? height)
    {
        var parameters = new OperationParameters();
        parameters.Set("pageIndex", pageIndex);
        if (width.HasValue) parameters.Set("width", width.Value);
        if (height.HasValue) parameters.Set("height", height.Value);
        return parameters;
    }

    /// <summary>
    ///     Builds parameters for the get page details operation.
    /// </summary>
    /// <param name="pageIndex">The page index (1-based) to get details for.</param>
    /// <returns>OperationParameters configured for getting page details.</returns>
    private static OperationParameters BuildGetDetailsParameters(int pageIndex)
    {
        return BuildDeleteParameters(pageIndex);
    }
}
