using System.ComponentModel;
using Aspose.Cells;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Session;
using ModelContextProtocol.Server;

namespace AsposeMcpServer.Tools.Excel;

/// <summary>
///     Unified tool for managing Excel filters (auto filter, custom filter, get filter status).
///     Merges: ExcelAutoFilterTool, ExcelGetFilterStatusTool.
/// </summary>
[McpServerToolType]
public class ExcelFilterTool
{
    /// <summary>
    ///     Handler registry for filter operations.
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
    ///     Initializes a new instance of the <see cref="ExcelFilterTool" /> class.
    /// </summary>
    /// <param name="sessionManager">Optional session manager for in-memory document editing.</param>
    /// <param name="identityAccessor">Optional session identity accessor for session isolation.</param>
    public ExcelFilterTool(DocumentSessionManager? sessionManager = null,
        ISessionIdentityAccessor? identityAccessor = null)
    {
        _sessionManager = sessionManager;
        _identityAccessor = identityAccessor;
        _handlerRegistry = HandlerRegistry<Workbook>.CreateFromNamespace("AsposeMcpServer.Handlers.Excel.Filter");
    }

    /// <summary>
    ///     Executes an Excel filter operation (apply, remove, filter, or get_status).
    /// </summary>
    /// <param name="operation">The operation to perform: apply, remove, filter, or get_status.</param>
    /// <param name="path">Excel file path (required if no sessionId).</param>
    /// <param name="sessionId">Session ID for in-memory editing.</param>
    /// <param name="outputPath">Output file path (file mode only).</param>
    /// <param name="sheetIndex">Sheet index (0-based).</param>
    /// <param name="range">Cell range to apply filter (e.g., 'A1:C10', required for apply/filter).</param>
    /// <param name="columnIndex">Column index within filter range to apply criteria (0-based, required for filter).</param>
    /// <param name="criteria">Filter criteria value (required for filter operation).</param>
    /// <param name="filterOperator">Filter operator for custom filter.</param>
    /// <returns>A message indicating the result of the operation, or JSON data for get_status operation.</returns>
    /// <exception cref="ArgumentException">Thrown when required parameters are missing or the operation is unknown.</exception>
    [McpServerTool(Name = "excel_filter")]
    [Description(@"Manage Excel filters. Supports 4 operations: apply, remove, filter, get_status.

Usage examples:
- Apply auto filter: excel_filter(operation='apply', path='book.xlsx', range='A1:C10')
- Remove filter: excel_filter(operation='remove', path='book.xlsx')
- Filter by value: excel_filter(operation='filter', path='book.xlsx', range='A1:C10', columnIndex=0, criteria='Completed')
- Filter by custom: excel_filter(operation='filter', path='book.xlsx', range='A1:C10', columnIndex=1, filterOperator='GreaterThan', criteria='100')
- Get filter status: excel_filter(operation='get_status', path='book.xlsx')")]
    public string Execute(
        [Description(@"Operation to perform.
- 'apply': Apply auto filter dropdown buttons (required params: path, range)
- 'remove': Remove auto filter completely (required params: path)
- 'filter': Apply filter criteria to a column (required params: path, range, columnIndex, criteria)
- 'get_status': Get filter status with details (required params: path)")]
        string operation,
        [Description("Excel file path (required if no sessionId)")]
        string? path = null,
        [Description("Session ID for in-memory editing")]
        string? sessionId = null,
        [Description("Output file path (file mode only)")]
        string? outputPath = null,
        [Description("Sheet index (0-based, default: 0)")]
        int sheetIndex = 0,
        [Description("Cell range to apply filter (e.g., 'A1:C10', required for apply/filter)")]
        string? range = null,
        [Description("Column index within filter range to apply criteria (0-based, required for filter)")]
        int columnIndex = 0,
        [Description("Filter criteria value (required for filter operation)")]
        string? criteria = null,
        [Description("Filter operator for custom filter (optional, default: 'Equal')")]
        string filterOperator = "Equal")
    {
        using var ctx = DocumentContext<Workbook>.Create(_sessionManager, sessionId, path, _identityAccessor);

        var parameters = BuildParameters(operation, sheetIndex, range, columnIndex, criteria, filterOperator);

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

        if (operation.ToLowerInvariant() == "get_status")
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
        int sheetIndex,
        string? range,
        int columnIndex,
        string? criteria,
        string filterOperator)
    {
        var parameters = new OperationParameters();
        parameters.Set("sheetIndex", sheetIndex);

        switch (operation.ToLowerInvariant())
        {
            case "apply":
                if (range != null) parameters.Set("range", range);
                break;

            case "remove":
                break;

            case "filter":
                if (range != null) parameters.Set("range", range);
                parameters.Set("columnIndex", columnIndex);
                if (criteria != null) parameters.Set("criteria", criteria);
                parameters.Set("filterOperator", filterOperator);
                break;

            case "get_status":
                break;
        }

        return parameters;
    }
}
