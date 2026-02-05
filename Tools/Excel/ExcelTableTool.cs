using System.ComponentModel;
using Aspose.Cells;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Session;
using AsposeMcpServer.Helpers;
using ModelContextProtocol.Server;

namespace AsposeMcpServer.Tools.Excel;

/// <summary>
///     Unified tool for managing Excel tables/ListObjects (create, get, delete, set_style, add_total_row,
///     convert_to_range).
/// </summary>
[ToolHandlerMapping("AsposeMcpServer.Handlers.Excel.Table")]
[McpServerToolType]
public class ExcelTableTool
{
    /// <summary>
    ///     Handler registry for table operations.
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
    ///     Initializes a new instance of the <see cref="ExcelTableTool" /> class.
    /// </summary>
    /// <param name="sessionManager">Optional session manager for in-memory document editing.</param>
    /// <param name="identityAccessor">Optional session identity accessor for session isolation.</param>
    public ExcelTableTool(DocumentSessionManager? sessionManager = null,
        ISessionIdentityAccessor? identityAccessor = null)
    {
        _sessionManager = sessionManager;
        _identityAccessor = identityAccessor;
        _handlerRegistry = HandlerRegistry<Workbook>.CreateFromNamespace("AsposeMcpServer.Handlers.Excel.Table");
    }

    /// <summary>
    ///     Executes an Excel table operation (create, get, delete, set_style, add_total_row, convert_to_range).
    /// </summary>
    /// <param name="operation">The operation to perform.</param>
    /// <param name="path">Excel file path (required if no sessionId).</param>
    /// <param name="sessionId">Session ID for in-memory editing.</param>
    /// <param name="outputPath">Output file path (file mode only).</param>
    /// <param name="sheetIndex">Sheet index (0-based).</param>
    /// <param name="range">Cell range for table creation (e.g., 'A1:D10').</param>
    /// <param name="hasHeaders">Whether the range has headers (default: true, for create).</param>
    /// <param name="name">Table name (optional, for create).</param>
    /// <param name="tableIndex">Table index (0-based, for delete/set_style/add_total_row/convert_to_range).</param>
    /// <param name="styleName">Table style name (for set_style, e.g., 'TableStyleMedium9').</param>
    /// <param name="keepData">Whether to keep data when deleting (default: true).</param>
    /// <param name="columnIndex">Column index for total function (for add_total_row).</param>
    /// <param name="totalFunction">Total function: sum, count, average, max, min, none (for add_total_row).</param>
    /// <returns>A message or data indicating the result of the operation.</returns>
    /// <exception cref="ArgumentException">Thrown when required parameters are missing or the operation is unknown.</exception>
    [McpServerTool(
        Name = "excel_table",
        Title = "Excel Table Operations",
        Destructive = true,
        Idempotent = false,
        OpenWorld = false,
        ReadOnly = false,
        UseStructuredContent = true)]
    [Description(
        @"Manage Excel tables (ListObjects). Supports 6 operations: create, get, delete, set_style, add_total_row, convert_to_range.

Usage examples:
- Create table: excel_table(operation='create', path='book.xlsx', range='A1:D10')
- Create named table: excel_table(operation='create', path='book.xlsx', range='A1:D10', name='SalesData')
- Get tables: excel_table(operation='get', path='book.xlsx')
- Delete table: excel_table(operation='delete', path='book.xlsx', tableIndex=0)
- Set style: excel_table(operation='set_style', path='book.xlsx', tableIndex=0, styleName='TableStyleMedium9')
- Add total row: excel_table(operation='add_total_row', path='book.xlsx', tableIndex=0, columnIndex=2, totalFunction='sum')
- Convert to range: excel_table(operation='convert_to_range', path='book.xlsx', tableIndex=0)")]
    public object Execute(
        [Description(@"Operation to perform.
- 'create': Create a table from a cell range (required params: range)
- 'get': Get table(s) information (optional: tableIndex)
- 'delete': Delete a table (required params: tableIndex)
- 'set_style': Set table style (required params: tableIndex, styleName)
- 'add_total_row': Add/configure total row (required params: tableIndex)
- 'convert_to_range': Convert table to normal range (required params: tableIndex)")]
        string operation,
        [Description("Excel file path (required if no sessionId)")]
        string? path = null,
        [Description("Session ID for in-memory editing")]
        string? sessionId = null,
        [Description("Output file path (file mode only)")]
        string? outputPath = null,
        [Description("Sheet index (0-based, default: 0)")]
        int sheetIndex = 0,
        [Description("Cell range for table creation (e.g., 'A1:D10', required for create)")]
        string? range = null,
        [Description("Whether the range has headers (default: true, for create)")]
        bool hasHeaders = true,
        [Description("Table name (optional, for create)")]
        string? name = null,
        [Description("Table index (0-based, for delete/set_style/add_total_row/convert_to_range)")]
        int? tableIndex = null,
        [Description("Table style name (for set_style, e.g., 'TableStyleMedium9')")]
        string? styleName = null,
        [Description("Whether to keep data when deleting (default: true)")]
        bool keepData = true,
        [Description("Column index for total function (for add_total_row)")]
        int? columnIndex = null,
        [Description("Total function: sum, count, average, max, min, none (for add_total_row)")]
        string? totalFunction = null)
    {
        using var ctx = DocumentContext<Workbook>.Create(_sessionManager, sessionId, path, _identityAccessor);

        var parameters = BuildParameters(operation, sheetIndex, range, hasHeaders, name, tableIndex, styleName,
            keepData, columnIndex, totalFunction);

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
        string? range,
        bool hasHeaders,
        string? name,
        int? tableIndex,
        string? styleName,
        bool keepData,
        int? columnIndex,
        string? totalFunction)
    {
        var parameters = new OperationParameters();
        parameters.Set("sheetIndex", sheetIndex);

        return operation.ToLowerInvariant() switch
        {
            "create" => BuildCreateParameters(parameters, range, hasHeaders, name),
            "get" => BuildGetParameters(parameters, tableIndex),
            "delete" => BuildDeleteParameters(parameters, tableIndex, keepData),
            "set_style" => BuildSetStyleParameters(parameters, tableIndex, styleName),
            "add_total_row" => BuildAddTotalRowParameters(parameters, tableIndex, columnIndex, totalFunction),
            "convert_to_range" => BuildConvertToRangeParameters(parameters, tableIndex),
            _ => parameters
        };
    }

    /// <summary>
    ///     Builds parameters for the create operation.
    /// </summary>
    /// <param name="parameters">Base parameters with sheet index.</param>
    /// <param name="range">The cell range.</param>
    /// <param name="hasHeaders">Whether the range has headers.</param>
    /// <param name="name">Optional table name.</param>
    /// <returns>OperationParameters configured for creating a table.</returns>
    private static OperationParameters BuildCreateParameters(OperationParameters parameters, string? range,
        bool hasHeaders, string? name)
    {
        if (range != null) parameters.Set("range", range);
        parameters.Set("hasHeaders", hasHeaders);
        if (name != null) parameters.Set("name", name);
        return parameters;
    }

    /// <summary>
    ///     Builds parameters for the get operation.
    /// </summary>
    /// <param name="parameters">Base parameters with sheet index.</param>
    /// <param name="tableIndex">Optional specific table index.</param>
    /// <returns>OperationParameters configured for getting tables.</returns>
    private static OperationParameters BuildGetParameters(OperationParameters parameters, int? tableIndex)
    {
        if (tableIndex.HasValue) parameters.Set("tableIndex", tableIndex.Value);
        return parameters;
    }

    /// <summary>
    ///     Builds parameters for the delete operation.
    /// </summary>
    /// <param name="parameters">Base parameters with sheet index.</param>
    /// <param name="tableIndex">The table index.</param>
    /// <param name="keepData">Whether to keep data.</param>
    /// <returns>OperationParameters configured for deleting a table.</returns>
    private static OperationParameters BuildDeleteParameters(OperationParameters parameters, int? tableIndex,
        bool keepData)
    {
        if (tableIndex.HasValue) parameters.Set("tableIndex", tableIndex.Value);
        parameters.Set("keepData", keepData);
        return parameters;
    }

    /// <summary>
    ///     Builds parameters for the set_style operation.
    /// </summary>
    /// <param name="parameters">Base parameters with sheet index.</param>
    /// <param name="tableIndex">The table index.</param>
    /// <param name="styleName">The style name.</param>
    /// <returns>OperationParameters configured for setting table style.</returns>
    private static OperationParameters BuildSetStyleParameters(OperationParameters parameters, int? tableIndex,
        string? styleName)
    {
        if (tableIndex.HasValue) parameters.Set("tableIndex", tableIndex.Value);
        if (styleName != null) parameters.Set("styleName", styleName);
        return parameters;
    }

    /// <summary>
    ///     Builds parameters for the add_total_row operation.
    /// </summary>
    /// <param name="parameters">Base parameters with sheet index.</param>
    /// <param name="tableIndex">The table index.</param>
    /// <param name="columnIndex">The column index for total function.</param>
    /// <param name="totalFunction">The total function name.</param>
    /// <returns>OperationParameters configured for adding a total row.</returns>
    private static OperationParameters BuildAddTotalRowParameters(OperationParameters parameters, int? tableIndex,
        int? columnIndex, string? totalFunction)
    {
        if (tableIndex.HasValue) parameters.Set("tableIndex", tableIndex.Value);
        if (columnIndex.HasValue) parameters.Set("columnIndex", columnIndex.Value);
        if (totalFunction != null) parameters.Set("totalFunction", totalFunction);
        return parameters;
    }

    /// <summary>
    ///     Builds parameters for the convert_to_range operation.
    /// </summary>
    /// <param name="parameters">Base parameters with sheet index.</param>
    /// <param name="tableIndex">The table index.</param>
    /// <returns>OperationParameters configured for converting table to range.</returns>
    private static OperationParameters BuildConvertToRangeParameters(OperationParameters parameters, int? tableIndex)
    {
        if (tableIndex.HasValue) parameters.Set("tableIndex", tableIndex.Value);
        return parameters;
    }
}
