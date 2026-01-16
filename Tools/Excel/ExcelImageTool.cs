using System.ComponentModel;
using Aspose.Cells;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Session;
using ModelContextProtocol.Server;

namespace AsposeMcpServer.Tools.Excel;

/// <summary>
///     Unified tool for managing Excel images (add, delete, get, extract).
/// </summary>
[McpServerToolType]
public class ExcelImageTool
{
    /// <summary>
    ///     Handler registry for image operations.
    /// </summary>
    private readonly HandlerRegistry<Workbook> _handlerRegistry;

    /// <summary>
    ///     Session identity accessor for session isolation support.
    /// </summary>
    private readonly ISessionIdentityAccessor? _identityAccessor;

    /// <summary>
    ///     Document session manager for in-memory editing support.
    /// </summary>
    private readonly DocumentSessionManager? _sessionManager;

    /// <summary>
    ///     Initializes a new instance of the <see cref="ExcelImageTool" /> class.
    /// </summary>
    /// <param name="sessionManager">Optional session manager for in-memory document editing.</param>
    /// <param name="identityAccessor">Optional session identity accessor for session isolation.</param>
    public ExcelImageTool(DocumentSessionManager? sessionManager = null,
        ISessionIdentityAccessor? identityAccessor = null)
    {
        _sessionManager = sessionManager;
        _identityAccessor = identityAccessor;
        _handlerRegistry = HandlerRegistry<Workbook>.CreateFromNamespace("AsposeMcpServer.Handlers.Excel.Image");
    }

    /// <summary>
    ///     Executes an Excel image operation (add, delete, get, extract).
    /// </summary>
    /// <param name="operation">The operation to perform: add, delete, get, extract.</param>
    /// <param name="path">Excel file path (required if no sessionId).</param>
    /// <param name="sessionId">Session ID for in-memory editing.</param>
    /// <param name="outputPath">Output file path (file mode only).</param>
    /// <param name="sheetIndex">Sheet index (0-based, default: 0).</param>
    /// <param name="imagePath">
    ///     Path to the image file. Supported formats: png, jpg, jpeg, gif, bmp, tiff, emf, wmf (required
    ///     for add).
    /// </param>
    /// <param name="cell">Top-left cell reference (e.g., 'A1', required for add).</param>
    /// <param name="width">Image width in pixels (optional for add).</param>
    /// <param name="height">Image height in pixels (optional for add).</param>
    /// <param name="keepAspectRatio">Keep aspect ratio when resizing (optional for add, default: true).</param>
    /// <param name="imageIndex">Image index (0-based, required for delete/extract).</param>
    /// <param name="exportPath">Path to export the extracted image (required for extract).</param>
    /// <returns>A message indicating the result of the operation, or JSON data for get operations.</returns>
    /// <exception cref="ArgumentException">Thrown when required parameters are missing or the operation is unknown.</exception>
    [McpServerTool(Name = "excel_image")]
    [Description(@"Manage Excel images. Supports 4 operations: add, delete, get, extract.

Usage examples:
- Add image: excel_image(operation='add', path='book.xlsx', imagePath='image.png', cell='A1', width=200, height=150)
- Add image with aspect ratio: excel_image(operation='add', path='book.xlsx', imagePath='image.png', cell='A1', width=200, keepAspectRatio=true)
- Delete image: excel_image(operation='delete', path='book.xlsx', imageIndex=0)
- Get images: excel_image(operation='get', path='book.xlsx')
- Extract image: excel_image(operation='extract', path='book.xlsx', imageIndex=0, exportPath='extracted.png')

Note: When deleting images, the indices of remaining images will be re-ordered.")]
    public string Execute( // NOSONAR S107 - MCP protocol requires multiple parameters
        [Description("Operation: add, delete, get, extract")]
        string operation,
        [Description("Excel file path (required if no sessionId)")]
        string? path = null,
        [Description("Session ID for in-memory editing")]
        string? sessionId = null,
        [Description("Output file path (file mode only)")]
        string? outputPath = null,
        [Description("Sheet index (0-based, default: 0)")]
        int sheetIndex = 0,
        [Description(
            "Path to the image file. Supported formats: png, jpg, jpeg, gif, bmp, tiff, emf, wmf (required for add)")]
        string? imagePath = null,
        [Description("Top-left cell reference (e.g., 'A1', required for add)")]
        string? cell = null,
        [Description("Image width in pixels (optional for add)")]
        int? width = null,
        [Description("Image height in pixels (optional for add)")]
        int? height = null,
        [Description(
            "Keep aspect ratio when resizing. If true and only width or height is specified, the other dimension is calculated proportionally (optional for add, default: true)")]
        bool keepAspectRatio = true,
        [Description("Image index (0-based, required for delete/extract). Note: indices are re-ordered after deletion")]
        int? imageIndex = null,
        [Description(
            "Path to export the extracted image (required for extract). Format determined by file extension (png, jpg, gif, bmp, tiff)")]
        string? exportPath = null)
    {
        using var ctx = DocumentContext<Workbook>.Create(_sessionManager, sessionId, path, _identityAccessor);

        var parameters = BuildParameters(operation, sheetIndex, imagePath, cell, width, height, keepAspectRatio,
            imageIndex, exportPath);

        var handler = _handlerRegistry.GetHandler(operation);

        var operationContext = new OperationContext<Workbook>
        {
            Document = ctx.Document,
            SessionManager = _sessionManager,
            IdentityAccessor = _identityAccessor,
            SessionId = sessionId,
            SourcePath = path,
            OutputPath = outputPath
        };

        var result = handler.Execute(operationContext, parameters);

        if (operation.ToLowerInvariant() is "get" or "extract")
            return result;

        if (operationContext.IsModified)
            ctx.Save(outputPath);

        return $"{result}\n{ctx.GetOutputMessage(outputPath)}";
    }

    /// <summary>
    ///     Builds OperationParameters from method parameters using strategy pattern.
    ///     Parameters are documented on the Execute method.
    /// </summary>
    /// <returns>OperationParameters configured with all input values.</returns>
    private static OperationParameters BuildParameters( // NOSONAR S107
        string operation,
        int sheetIndex,
        string? imagePath,
        string? cell,
        int? width,
        int? height,
        bool keepAspectRatio,
        int? imageIndex,
        string? exportPath)
    {
        var parameters = new OperationParameters();
        parameters.Set("sheetIndex", sheetIndex);

        return operation.ToLowerInvariant() switch
        {
            "add" => BuildAddParameters(parameters, imagePath, cell, width, height, keepAspectRatio),
            "delete" => BuildDeleteParameters(parameters, imageIndex),
            "get" => parameters,
            "extract" => BuildExtractParameters(parameters, imageIndex, exportPath),
            _ => parameters
        };
    }

    /// <summary>
    ///     Builds parameters for the add image operation.
    /// </summary>
    /// <param name="parameters">Base parameters with sheet index.</param>
    /// <param name="imagePath">The path to the image file.</param>
    /// <param name="cell">The top-left cell reference.</param>
    /// <param name="width">The image width in pixels.</param>
    /// <param name="height">The image height in pixels.</param>
    /// <param name="keepAspectRatio">Whether to keep aspect ratio when resizing.</param>
    /// <returns>OperationParameters configured for adding image.</returns>
    private static OperationParameters BuildAddParameters(OperationParameters parameters, string? imagePath,
        string? cell, int? width, int? height, bool keepAspectRatio)
    {
        if (imagePath != null) parameters.Set("imagePath", imagePath);
        if (cell != null) parameters.Set("cell", cell);
        if (width.HasValue) parameters.Set("width", width.Value);
        if (height.HasValue) parameters.Set("height", height.Value);
        parameters.Set("keepAspectRatio", keepAspectRatio);
        return parameters;
    }

    /// <summary>
    ///     Builds parameters for the delete image operation.
    /// </summary>
    /// <param name="parameters">Base parameters with sheet index.</param>
    /// <param name="imageIndex">The index of image to delete.</param>
    /// <returns>OperationParameters configured for deleting image.</returns>
    private static OperationParameters BuildDeleteParameters(OperationParameters parameters, int? imageIndex)
    {
        if (imageIndex.HasValue) parameters.Set("imageIndex", imageIndex.Value);
        return parameters;
    }

    /// <summary>
    ///     Builds parameters for the extract image operation.
    /// </summary>
    /// <param name="parameters">Base parameters with sheet index.</param>
    /// <param name="imageIndex">The index of image to extract.</param>
    /// <param name="exportPath">The path to export the extracted image.</param>
    /// <returns>OperationParameters configured for extracting image.</returns>
    private static OperationParameters BuildExtractParameters(OperationParameters parameters, int? imageIndex,
        string? exportPath)
    {
        if (imageIndex.HasValue) parameters.Set("imageIndex", imageIndex.Value);
        if (exportPath != null) parameters.Set("exportPath", exportPath);
        return parameters;
    }
}
