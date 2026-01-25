using System.ComponentModel;
using Aspose.Cells;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Session;
using AsposeMcpServer.Helpers;
using ModelContextProtocol.Server;

namespace AsposeMcpServer.Tools.Excel;

/// <summary>
///     Unified tool for managing Excel freeze panes (freeze/unfreeze/get).
/// </summary>
[ToolHandlerMapping("AsposeMcpServer.Handlers.Excel.FreezePanes")]
[McpServerToolType]
public class ExcelFreezePanesTool
{
    /// <summary>
    ///     Handler registry for freeze panes operations.
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
    ///     Initializes a new instance of the <see cref="ExcelFreezePanesTool" /> class.
    /// </summary>
    /// <param name="sessionManager">Optional session manager for in-memory document editing.</param>
    /// <param name="identityAccessor">Optional session identity accessor for session isolation.</param>
    public ExcelFreezePanesTool(DocumentSessionManager? sessionManager = null,
        ISessionIdentityAccessor? identityAccessor = null)
    {
        _sessionManager = sessionManager;
        _identityAccessor = identityAccessor;
        _handlerRegistry = HandlerRegistry<Workbook>.CreateFromNamespace("AsposeMcpServer.Handlers.Excel.FreezePanes");
    }

    /// <summary>
    ///     Executes an Excel freeze panes operation (freeze, unfreeze, get).
    /// </summary>
    /// <param name="operation">The operation to perform: freeze, unfreeze, get.</param>
    /// <param name="path">Excel file path (required if no sessionId).</param>
    /// <param name="sessionId">Session ID for in-memory editing.</param>
    /// <param name="outputPath">Output file path (file mode only).</param>
    /// <param name="sheetIndex">Sheet index (0-based, default: 0).</param>
    /// <param name="row">Number of rows to freeze from top (0-based, required for freeze).</param>
    /// <param name="column">Number of columns to freeze from left (0-based, required for freeze).</param>
    /// <returns>A message indicating the result of the operation, or JSON data for get operations.</returns>
    /// <exception cref="ArgumentException">Thrown when required parameters are missing or the operation is unknown.</exception>
    [McpServerTool(
        Name = "excel_freeze_panes",
        Title = "Excel Freeze Panes Operations",
        Destructive = false,
        Idempotent = true,
        OpenWorld = false,
        ReadOnly = false,
        UseStructuredContent = true)]
    [Description(@"Manage Excel freeze panes. Supports 3 operations: freeze, unfreeze, get.

Usage examples:
- Freeze panes: excel_freeze_panes(operation='freeze', path='book.xlsx', row=1, column=1)
- Unfreeze panes: excel_freeze_panes(operation='unfreeze', path='book.xlsx')
- Get freeze status: excel_freeze_panes(operation='get', path='book.xlsx')")]
    public object Execute(
        [Description(@"Operation to perform.
- 'freeze': Freeze panes at specified row and column (required params: path, row, column)
- 'unfreeze': Remove freeze panes (required params: path)
- 'get': Get current freeze panes status (required params: path)")]
        string operation,
        [Description("Excel file path (required if no sessionId)")]
        string? path = null,
        [Description("Session ID for in-memory editing")]
        string? sessionId = null,
        [Description("Output file path (file mode only)")]
        string? outputPath = null,
        [Description("Sheet index (0-based, default: 0)")]
        int sheetIndex = 0,
        [Description("Number of rows to freeze from top (0-based, required for freeze)")]
        int row = 0,
        [Description("Number of columns to freeze from left (0-based, required for freeze)")]
        int column = 0)
    {
        using var ctx = DocumentContext<Workbook>.Create(_sessionManager, sessionId, path, _identityAccessor);

        var parameters = BuildParameters(operation, sheetIndex, row, column);

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
        int sheetIndex,
        int row,
        int column)
    {
        return operation.ToLowerInvariant() switch
        {
            "freeze" => BuildFreezeParameters(sheetIndex, row, column),
            "unfreeze" or "get" => BuildBaseParameters(sheetIndex),
            _ => new OperationParameters()
        };
    }

    /// <summary>
    ///     Builds parameters for the freeze panes operation.
    /// </summary>
    /// <param name="sheetIndex">The sheet index (0-based).</param>
    /// <param name="row">Number of rows to freeze from top.</param>
    /// <param name="column">Number of columns to freeze from left.</param>
    /// <returns>OperationParameters configured for freezing panes.</returns>
    private static OperationParameters BuildFreezeParameters(int sheetIndex, int row, int column)
    {
        var parameters = new OperationParameters();
        parameters.Set("sheetIndex", sheetIndex);
        parameters.Set("row", row);
        parameters.Set("column", column);
        return parameters;
    }

    /// <summary>
    ///     Builds base parameters containing only the sheet index.
    /// </summary>
    /// <param name="sheetIndex">The sheet index (0-based).</param>
    /// <returns>OperationParameters with sheet index set.</returns>
    private static OperationParameters BuildBaseParameters(int sheetIndex)
    {
        var parameters = new OperationParameters();
        parameters.Set("sheetIndex", sheetIndex);
        return parameters;
    }
}
