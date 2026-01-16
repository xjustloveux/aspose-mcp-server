using System.ComponentModel;
using Aspose.Cells;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Session;
using ModelContextProtocol.Server;

namespace AsposeMcpServer.Tools.Excel;

/// <summary>
///     Unified tool for managing Excel sheets (add, delete, get, rename, move, copy, hide)
/// </summary>
[McpServerToolType]
public class ExcelSheetTool
{
    /// <summary>
    ///     Handler registry for sheet operations.
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
    ///     Initializes a new instance of the <see cref="ExcelSheetTool" /> class.
    /// </summary>
    /// <param name="sessionManager">Optional session manager for in-memory document editing.</param>
    /// <param name="identityAccessor">Optional session identity accessor for session isolation.</param>
    public ExcelSheetTool(DocumentSessionManager? sessionManager = null,
        ISessionIdentityAccessor? identityAccessor = null)
    {
        _sessionManager = sessionManager;
        _identityAccessor = identityAccessor;
        _handlerRegistry = HandlerRegistry<Workbook>.CreateFromNamespace("AsposeMcpServer.Handlers.Excel.Sheet");
    }

    /// <summary>
    ///     Executes an Excel sheet operation (add, delete, get, rename, move, copy, or hide).
    /// </summary>
    /// <param name="operation">The operation to perform: add, delete, get, rename, move, copy, or hide.</param>
    /// <param name="path">Excel file path (required if no sessionId).</param>
    /// <param name="sessionId">Session ID for in-memory editing.</param>
    /// <param name="outputPath">Output file path (file mode only).</param>
    /// <param name="sheetIndex">Sheet index (0-based, required for delete/rename/move/copy/hide).</param>
    /// <param name="sheetName">Name of the sheet (required for add operation).</param>
    /// <param name="newName">New name for the sheet (required for rename, max 31 characters).</param>
    /// <param name="insertAt">Position to insert the sheet (0-based, optional for add/move).</param>
    /// <param name="targetIndex">Target index for move/copy operation (0-based).</param>
    /// <param name="copyToPath">Target file path for copy operation (optional).</param>
    /// <returns>A message indicating the result of the operation, or JSON data for get operation.</returns>
    /// <exception cref="ArgumentException">Thrown when required parameters are missing or the operation is unknown.</exception>
    /// <exception cref="InvalidOperationException">Thrown when attempting to delete the last worksheet.</exception>
    [McpServerTool(Name = "excel_sheet")]
    [Description(@"Manage Excel sheets. Supports 7 operations: add, delete, get, rename, move, copy, hide.

Usage examples:
- Add sheet: excel_sheet(operation='add', path='book.xlsx', sheetName='New Sheet')
- Delete sheet: excel_sheet(operation='delete', path='book.xlsx', sheetIndex=1)
- Get sheets: excel_sheet(operation='get', path='book.xlsx')
- Rename sheet: excel_sheet(operation='rename', path='book.xlsx', sheetIndex=0, newName='Renamed')
- Move sheet: excel_sheet(operation='move', path='book.xlsx', sheetIndex=0, insertAt=2)
- Copy sheet: excel_sheet(operation='copy', path='book.xlsx', sheetIndex=0, newName='Copy')
- Hide sheet: excel_sheet(operation='hide', path='book.xlsx', sheetIndex=1)")]
    public string Execute(
        [Description("Operation to perform: add, delete, get, rename, move, copy, hide")]
        string operation,
        [Description("Excel file path (required if no sessionId)")]
        string? path = null,
        [Description("Session ID for in-memory editing")]
        string? sessionId = null,
        [Description("Output file path (file mode only)")]
        string? outputPath = null,
        [Description("Sheet index (0-based, required for delete/rename/move/copy/hide)")]
        int sheetIndex = 0,
        [Description("Name of the sheet (required for add operation)")]
        string? sheetName = null,
        [Description("New name for the sheet (required for rename, max 31 characters)")]
        string? newName = null,
        [Description("Position to insert the sheet (0-based, optional for add/move)")]
        int? insertAt = null,
        [Description("Target index for move/copy operation (0-based)")]
        int? targetIndex = null,
        [Description("Target file path for copy operation (optional)")]
        string? copyToPath = null)
    {
        using var ctx = DocumentContext<Workbook>.Create(_sessionManager, sessionId, path, _identityAccessor);

        var parameters = BuildParameters(operation, sheetIndex, sheetName, newName, insertAt, targetIndex, copyToPath);

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

        if (string.Equals(operation, "get", StringComparison.OrdinalIgnoreCase))
            return result;

        if (operationContext.IsModified)
            ctx.Save(outputPath);

        return $"{result}\n{ctx.GetOutputMessage(outputPath)}";
    }

    /// <summary>
    ///     Builds OperationParameters from method parameters using strategy pattern.
    /// </summary>
    private static OperationParameters BuildParameters(
        string operation,
        int sheetIndex,
        string? sheetName,
        string? newName,
        int? insertAt,
        int? targetIndex,
        string? copyToPath)
    {
        var parameters = new OperationParameters();
        parameters.Set("sheetIndex", sheetIndex);

        return operation.ToLowerInvariant() switch
        {
            "add" => BuildAddParameters(parameters, sheetName, insertAt),
            "delete" or "get" or "hide" => parameters,
            "rename" => BuildRenameParameters(parameters, newName),
            "move" => BuildMoveParameters(parameters, targetIndex, insertAt),
            "copy" => BuildCopyParameters(parameters, targetIndex, copyToPath),
            _ => parameters
        };
    }

    /// <summary>
    ///     Builds parameters for the add sheet operation.
    /// </summary>
    /// <param name="parameters">The base operation parameters.</param>
    /// <param name="sheetName">The name for the new sheet.</param>
    /// <param name="insertAt">The position to insert the sheet (0-based).</param>
    /// <returns>OperationParameters configured for the add operation.</returns>
    private static OperationParameters BuildAddParameters(OperationParameters parameters, string? sheetName,
        int? insertAt)
    {
        if (sheetName != null) parameters.Set("sheetName", sheetName);
        if (insertAt.HasValue) parameters.Set("insertAt", insertAt.Value);
        return parameters;
    }

    /// <summary>
    ///     Builds parameters for the rename sheet operation.
    /// </summary>
    /// <param name="parameters">The base operation parameters.</param>
    /// <param name="newName">The new name for the sheet (max 31 characters).</param>
    /// <returns>OperationParameters configured for the rename operation.</returns>
    private static OperationParameters BuildRenameParameters(OperationParameters parameters, string? newName)
    {
        if (newName != null) parameters.Set("newName", newName);
        return parameters;
    }

    /// <summary>
    ///     Builds parameters for the move sheet operation.
    /// </summary>
    /// <param name="parameters">The base operation parameters.</param>
    /// <param name="targetIndex">The target index for the move operation (0-based).</param>
    /// <param name="insertAt">The position to insert the sheet at (0-based).</param>
    /// <returns>OperationParameters configured for the move operation.</returns>
    private static OperationParameters BuildMoveParameters(OperationParameters parameters, int? targetIndex,
        int? insertAt)
    {
        if (targetIndex.HasValue) parameters.Set("targetIndex", targetIndex.Value);
        if (insertAt.HasValue) parameters.Set("insertAt", insertAt.Value);
        return parameters;
    }

    /// <summary>
    ///     Builds parameters for the copy sheet operation.
    /// </summary>
    /// <param name="parameters">The base operation parameters.</param>
    /// <param name="targetIndex">The target index for the copy operation (0-based).</param>
    /// <param name="copyToPath">The target file path for the copy operation.</param>
    /// <returns>OperationParameters configured for the copy operation.</returns>
    private static OperationParameters BuildCopyParameters(OperationParameters parameters, int? targetIndex,
        string? copyToPath)
    {
        if (targetIndex.HasValue) parameters.Set("targetIndex", targetIndex.Value);
        if (copyToPath != null) parameters.Set("copyToPath", copyToPath);
        return parameters;
    }
}
