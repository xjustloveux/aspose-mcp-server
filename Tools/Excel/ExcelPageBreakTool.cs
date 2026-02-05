using System.ComponentModel;
using Aspose.Cells;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Session;
using AsposeMcpServer.Helpers;
using ModelContextProtocol.Server;

namespace AsposeMcpServer.Tools.Excel;

/// <summary>
///     Unified tool for managing Excel page breaks (add_horizontal, add_vertical, remove, clear, get).
/// </summary>
[ToolHandlerMapping("AsposeMcpServer.Handlers.Excel.PageBreak")]
[McpServerToolType]
public class ExcelPageBreakTool
{
    /// <summary>
    ///     Handler registry for page break operations.
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
    ///     Initializes a new instance of the <see cref="ExcelPageBreakTool" /> class.
    /// </summary>
    /// <param name="sessionManager">Optional session manager for in-memory document editing.</param>
    /// <param name="identityAccessor">Optional session identity accessor for session isolation.</param>
    public ExcelPageBreakTool(DocumentSessionManager? sessionManager = null,
        ISessionIdentityAccessor? identityAccessor = null)
    {
        _sessionManager = sessionManager;
        _identityAccessor = identityAccessor;
        _handlerRegistry =
            HandlerRegistry<Workbook>.CreateFromNamespace("AsposeMcpServer.Handlers.Excel.PageBreak");
    }

    /// <summary>
    ///     Executes an Excel page break operation (add_horizontal, add_vertical, remove, clear, get).
    /// </summary>
    /// <param name="operation">The operation to perform.</param>
    /// <param name="path">Excel file path (required if no sessionId).</param>
    /// <param name="sessionId">Session ID for in-memory editing.</param>
    /// <param name="outputPath">Output file path (file mode only).</param>
    /// <param name="sheetIndex">Sheet index (0-based).</param>
    /// <param name="row">Row index (for add_horizontal).</param>
    /// <param name="column">Column index (for add_vertical).</param>
    /// <param name="breakType">Break type: horizontal, vertical, all (for remove/clear).</param>
    /// <param name="breakIndex">Break index (for remove).</param>
    /// <returns>A message or data indicating the result of the operation.</returns>
    /// <exception cref="ArgumentException">Thrown when required parameters are missing or the operation is unknown.</exception>
    [McpServerTool(
        Name = "excel_page_break",
        Title = "Excel Page Break Operations",
        Destructive = true,
        Idempotent = false,
        OpenWorld = false,
        ReadOnly = false,
        UseStructuredContent = true)]
    [Description(@"Manage Excel page breaks. Supports 5 operations: add_horizontal, add_vertical, remove, clear, get.

Usage examples:
- Add horizontal: excel_page_break(operation='add_horizontal', path='book.xlsx', row=10)
- Add vertical: excel_page_break(operation='add_vertical', path='book.xlsx', column=5)
- Remove: excel_page_break(operation='remove', path='book.xlsx', breakType='horizontal', breakIndex=0)
- Clear all: excel_page_break(operation='clear', path='book.xlsx')
- Get: excel_page_break(operation='get', path='book.xlsx')")]
    public object Execute(
        [Description(@"Operation to perform.
- 'add_horizontal': Add horizontal page break (required params: row)
- 'add_vertical': Add vertical page break (required params: column)
- 'remove': Remove a page break (required params: breakType, breakIndex)
- 'clear': Clear page breaks (optional: breakType defaults to 'all')
- 'get': Get all page breaks")]
        string operation,
        [Description("Excel file path (required if no sessionId)")]
        string? path = null,
        [Description("Session ID for in-memory editing")]
        string? sessionId = null,
        [Description("Output file path (file mode only)")]
        string? outputPath = null,
        [Description("Sheet index (0-based, default: 0)")]
        int sheetIndex = 0,
        [Description("Row index (for add_horizontal)")]
        int? row = null,
        [Description("Column index (for add_vertical)")]
        int? column = null,
        [Description("Break type: horizontal, vertical, all (for remove/clear)")]
        string? breakType = null,
        [Description("Break index (for remove)")]
        int? breakIndex = null)
    {
        using var ctx = DocumentContext<Workbook>.Create(_sessionManager, sessionId, path, _identityAccessor);

        var parameters = BuildParameters(operation, sheetIndex, row, column, breakType, breakIndex);

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
        int? row,
        int? column,
        string? breakType,
        int? breakIndex)
    {
        var parameters = new OperationParameters();
        parameters.Set("sheetIndex", sheetIndex);

        return operation.ToLowerInvariant() switch
        {
            "add_horizontal" => BuildAddHorizontalParameters(parameters, row),
            "add_vertical" => BuildAddVerticalParameters(parameters, column),
            "remove" => BuildRemoveParameters(parameters, breakType, breakIndex),
            "clear" => BuildClearParameters(parameters, breakType),
            _ => parameters
        };
    }

    /// <summary>
    ///     Builds parameters for the add_horizontal operation.
    /// </summary>
    private static OperationParameters BuildAddHorizontalParameters(OperationParameters parameters, int? row)
    {
        if (row.HasValue) parameters.Set("row", row.Value);
        return parameters;
    }

    /// <summary>
    ///     Builds parameters for the add_vertical operation.
    /// </summary>
    private static OperationParameters BuildAddVerticalParameters(OperationParameters parameters, int? column)
    {
        if (column.HasValue) parameters.Set("column", column.Value);
        return parameters;
    }

    /// <summary>
    ///     Builds parameters for the remove operation.
    /// </summary>
    private static OperationParameters BuildRemoveParameters(OperationParameters parameters, string? breakType,
        int? breakIndex)
    {
        if (breakType != null) parameters.Set("breakType", breakType);
        if (breakIndex.HasValue) parameters.Set("breakIndex", breakIndex.Value);
        return parameters;
    }

    /// <summary>
    ///     Builds parameters for the clear operation.
    /// </summary>
    private static OperationParameters BuildClearParameters(OperationParameters parameters, string? breakType)
    {
        if (breakType != null) parameters.Set("breakType", breakType);
        return parameters;
    }
}
