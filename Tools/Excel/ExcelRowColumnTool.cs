using System.ComponentModel;
using Aspose.Cells;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Session;
using AsposeMcpServer.Handlers.Excel.RowColumn;
using ModelContextProtocol.Server;

namespace AsposeMcpServer.Tools.Excel;

/// <summary>
///     Unified tool for managing Excel rows and columns (insert/delete rows, columns, cells)
/// </summary>
[McpServerToolType]
public class ExcelRowColumnTool
{
    /// <summary>
    ///     Handler registry for row/column operations.
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
    ///     Initializes a new instance of the <see cref="ExcelRowColumnTool" /> class.
    /// </summary>
    /// <param name="sessionManager">Optional session manager for in-memory document editing.</param>
    /// <param name="identityAccessor">Optional session identity accessor for session isolation.</param>
    public ExcelRowColumnTool(DocumentSessionManager? sessionManager = null,
        ISessionIdentityAccessor? identityAccessor = null)
    {
        _sessionManager = sessionManager;
        _identityAccessor = identityAccessor;
        _handlerRegistry = ExcelRowColumnHandlerRegistry.Create();
    }

    /// <summary>
    ///     Executes an Excel row/column operation (insert_row, delete_row, insert_column, delete_column, insert_cells,
    ///     delete_cells).
    /// </summary>
    /// <param name="operation">
    ///     The operation to perform: insert_row, delete_row, insert_column, delete_column, insert_cells,
    ///     delete_cells.
    /// </param>
    /// <param name="path">Excel file path (required if no sessionId).</param>
    /// <param name="sessionId">Session ID for in-memory editing.</param>
    /// <param name="outputPath">Output file path (file mode only).</param>
    /// <param name="sheetIndex">Sheet index (0-based, default: 0).</param>
    /// <param name="rowIndex">Row index (0-based, required for insert_row/delete_row).</param>
    /// <param name="columnIndex">Column index (0-based, required for insert_column/delete_column).</param>
    /// <param name="range">Cell range (e.g., 'A1:C5', required for insert_cells/delete_cells).</param>
    /// <param name="count">Number of rows/columns to insert/delete (default: 1).</param>
    /// <param name="shiftDirection">Shift direction: 'Right'/'Down' for insert_cells, 'Left'/'Up' for delete_cells.</param>
    /// <returns>A message indicating the result of the operation.</returns>
    /// <exception cref="ArgumentException">Thrown when required parameters are missing or the operation is unknown.</exception>
    [McpServerTool(Name = "excel_row_column")]
    [Description(
        @"Manage Excel rows and columns. Supports 6 operations: insert_row, delete_row, insert_column, delete_column, insert_cells, delete_cells.

Usage examples:
- Insert row: excel_row_column(operation='insert_row', path='book.xlsx', rowIndex=2, count=1)
- Delete row: excel_row_column(operation='delete_row', path='book.xlsx', rowIndex=2)
- Insert column: excel_row_column(operation='insert_column', path='book.xlsx', columnIndex=2, count=1)
- Delete column: excel_row_column(operation='delete_column', path='book.xlsx', columnIndex=2)
- Insert cells: excel_row_column(operation='insert_cells', path='book.xlsx', range='A1:C5', shiftDirection='Down')
- Delete cells: excel_row_column(operation='delete_cells', path='book.xlsx', range='A1:C5', shiftDirection='Up')")]
    public string Execute(
        [Description(
            "Operation to perform: insert_row, delete_row, insert_column, delete_column, insert_cells, delete_cells")]
        string operation,
        [Description("Excel file path (required if no sessionId)")]
        string? path = null,
        [Description("Session ID for in-memory editing")]
        string? sessionId = null,
        [Description("Output file path (file mode only)")]
        string? outputPath = null,
        [Description("Sheet index (0-based, default: 0)")]
        int sheetIndex = 0,
        [Description("Row index (0-based, required for insert_row/delete_row)")]
        int rowIndex = 0,
        [Description("Column index (0-based, required for insert_column/delete_column)")]
        int columnIndex = 0,
        [Description("Cell range (e.g., 'A1:C5', required for insert_cells/delete_cells)")]
        string? range = null,
        [Description("Number of rows/columns to insert/delete (default: 1)")]
        int count = 1,
        [Description("Shift direction: 'Right'/'Down' for insert_cells, 'Left'/'Up' for delete_cells")]
        string? shiftDirection = null)
    {
        var op = operation.ToLowerInvariant();
        if (op == "set_column_width")
            throw new ArgumentException(
                $"Operation 'set_column_width' is not supported by excel_row_column. Please use excel_view_settings operation instead. Example: excel_view_settings(operation='set_column_width', path='{path}', columnIndex=0, width=15)");

        using var ctx = DocumentContext<Workbook>.Create(_sessionManager, sessionId, path, _identityAccessor);

        var parameters = BuildParameters(operation, sheetIndex, rowIndex, columnIndex, range, count, shiftDirection);

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
        int rowIndex,
        int columnIndex,
        string? range,
        int count,
        string? shiftDirection)
    {
        var parameters = new OperationParameters();
        parameters.Set("sheetIndex", sheetIndex);

        switch (operation.ToLowerInvariant())
        {
            case "insert_row":
            case "delete_row":
                parameters.Set("rowIndex", rowIndex);
                parameters.Set("count", count);
                break;

            case "insert_column":
            case "delete_column":
                parameters.Set("columnIndex", columnIndex);
                parameters.Set("count", count);
                break;

            case "insert_cells":
            case "delete_cells":
                if (range != null) parameters.Set("range", range);
                if (shiftDirection != null) parameters.Set("shiftDirection", shiftDirection);
                break;
        }

        return parameters;
    }
}
