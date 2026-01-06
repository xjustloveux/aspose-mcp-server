using System.ComponentModel;
using Aspose.Cells;
using AsposeMcpServer.Core.Helpers;
using AsposeMcpServer.Core.Session;
using ModelContextProtocol.Server;

namespace AsposeMcpServer.Tools.Excel;

/// <summary>
///     Unified tool for managing Excel rows and columns (insert/delete rows, columns, cells)
///     Merges: ExcelInsertRowTool, ExcelDeleteRowTool, ExcelInsertColumnTool, ExcelDeleteColumnTool,
///     ExcelInsertCellsTool, ExcelDeleteCellsTool
/// </summary>
[McpServerToolType]
public class ExcelRowColumnTool
{
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
        using var ctx = DocumentContext<Workbook>.Create(_sessionManager, sessionId, path, _identityAccessor);

        return operation.ToLowerInvariant() switch
        {
            "insert_row" => InsertRow(ctx, outputPath, sheetIndex, rowIndex, count),
            "delete_row" => DeleteRow(ctx, outputPath, sheetIndex, rowIndex, count),
            "insert_column" => InsertColumn(ctx, outputPath, sheetIndex, columnIndex, count),
            "delete_column" => DeleteColumn(ctx, outputPath, sheetIndex, columnIndex, count),
            "insert_cells" => InsertCells(ctx, outputPath, sheetIndex, range, shiftDirection),
            "delete_cells" => DeleteCells(ctx, outputPath, sheetIndex, range, shiftDirection),
            "set_column_width" => throw new ArgumentException(
                $"Operation 'set_column_width' is not supported by excel_row_column. Please use excel_view_settings operation instead. Example: excel_view_settings(operation='set_column_width', path='{path}', columnIndex=0, width=15)"),
            _ => throw new ArgumentException($"Unknown operation: {operation}")
        };
    }

    /// <summary>
    ///     Inserts rows at the specified position.
    /// </summary>
    /// <param name="ctx">The document context containing the workbook.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="sheetIndex">The zero-based index of the worksheet.</param>
    /// <param name="rowIndex">The zero-based index where rows will be inserted.</param>
    /// <param name="count">The number of rows to insert.</param>
    /// <returns>A message indicating the result of the operation.</returns>
    private static string InsertRow(DocumentContext<Workbook> ctx, string? outputPath, int sheetIndex, int rowIndex,
        int count)
    {
        var workbook = ctx.Document;
        var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);

        worksheet.Cells.InsertRows(rowIndex, count);

        ctx.Save(outputPath);
        return $"Inserted {count} row(s) at row {rowIndex}. {ctx.GetOutputMessage(outputPath)}";
    }

    /// <summary>
    ///     Deletes rows starting from the specified position.
    /// </summary>
    /// <param name="ctx">The document context containing the workbook.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="sheetIndex">The zero-based index of the worksheet.</param>
    /// <param name="rowIndex">The zero-based index of the first row to delete.</param>
    /// <param name="count">The number of rows to delete.</param>
    /// <returns>A message indicating the result of the operation.</returns>
    private static string DeleteRow(DocumentContext<Workbook> ctx, string? outputPath, int sheetIndex, int rowIndex,
        int count)
    {
        var workbook = ctx.Document;
        var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);

        worksheet.Cells.DeleteRows(rowIndex, count);

        ctx.Save(outputPath);
        return $"Deleted {count} row(s) starting from row {rowIndex}. {ctx.GetOutputMessage(outputPath)}";
    }

    /// <summary>
    ///     Inserts columns at the specified position.
    /// </summary>
    /// <param name="ctx">The document context containing the workbook.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="sheetIndex">The zero-based index of the worksheet.</param>
    /// <param name="columnIndex">The zero-based index where columns will be inserted.</param>
    /// <param name="count">The number of columns to insert.</param>
    /// <returns>A message indicating the result of the operation.</returns>
    private static string InsertColumn(DocumentContext<Workbook> ctx, string? outputPath, int sheetIndex,
        int columnIndex, int count)
    {
        var workbook = ctx.Document;
        var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);

        worksheet.Cells.InsertColumns(columnIndex, count);

        ctx.Save(outputPath);
        return $"Inserted {count} column(s) at column {columnIndex}. {ctx.GetOutputMessage(outputPath)}";
    }

    /// <summary>
    ///     Deletes columns starting from the specified position.
    /// </summary>
    /// <param name="ctx">The document context containing the workbook.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="sheetIndex">The zero-based index of the worksheet.</param>
    /// <param name="columnIndex">The zero-based index of the first column to delete.</param>
    /// <param name="count">The number of columns to delete.</param>
    /// <returns>A message indicating the result of the operation.</returns>
    private static string DeleteColumn(DocumentContext<Workbook> ctx, string? outputPath, int sheetIndex,
        int columnIndex, int count)
    {
        var workbook = ctx.Document;
        var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);

        worksheet.Cells.DeleteColumns(columnIndex, count, true);

        ctx.Save(outputPath);
        return $"Deleted {count} column(s) starting from column {columnIndex}. {ctx.GetOutputMessage(outputPath)}";
    }

    /// <summary>
    ///     Inserts cells in a range and shifts existing cells.
    /// </summary>
    /// <param name="ctx">The document context containing the workbook.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="sheetIndex">The zero-based index of the worksheet.</param>
    /// <param name="range">The cell range where cells will be inserted (e.g., 'A1:C5').</param>
    /// <param name="shiftDirection">The direction to shift existing cells ('Right' or 'Down').</param>
    /// <returns>A message indicating the result of the operation.</returns>
    /// <exception cref="ArgumentException">Thrown when range or shiftDirection is null or empty.</exception>
    private static string InsertCells(DocumentContext<Workbook> ctx, string? outputPath, int sheetIndex, string? range,
        string? shiftDirection)
    {
        if (string.IsNullOrEmpty(range))
            throw new ArgumentException("range is required for insert_cells operation");
        if (string.IsNullOrEmpty(shiftDirection))
            throw new ArgumentException("shiftDirection is required for insert_cells operation");

        var workbook = ctx.Document;
        var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);

        var rangeObj = ExcelHelper.CreateRange(worksheet.Cells, range);
        var shiftType = string.Equals(shiftDirection, "right", StringComparison.OrdinalIgnoreCase)
            ? ShiftType.Right
            : ShiftType.Down;

        var cellArea = CellArea.CreateCellArea(
            rangeObj.FirstRow,
            rangeObj.FirstColumn,
            rangeObj.FirstRow + rangeObj.RowCount - 1,
            rangeObj.FirstColumn + rangeObj.ColumnCount - 1);

        worksheet.Cells.InsertRange(cellArea, shiftType);

        ctx.Save(outputPath);
        return $"Cells inserted in range {range}, shifted {shiftDirection}. {ctx.GetOutputMessage(outputPath)}";
    }

    /// <summary>
    ///     Deletes cells in a range and shifts remaining cells.
    /// </summary>
    /// <param name="ctx">The document context containing the workbook.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="sheetIndex">The zero-based index of the worksheet.</param>
    /// <param name="range">The cell range where cells will be deleted (e.g., 'A1:C5').</param>
    /// <param name="shiftDirection">The direction to shift remaining cells ('Left' or 'Up').</param>
    /// <returns>A message indicating the result of the operation.</returns>
    /// <exception cref="ArgumentException">Thrown when range or shiftDirection is null or empty.</exception>
    private static string DeleteCells(DocumentContext<Workbook> ctx, string? outputPath, int sheetIndex, string? range,
        string? shiftDirection)
    {
        if (string.IsNullOrEmpty(range))
            throw new ArgumentException("range is required for delete_cells operation");
        if (string.IsNullOrEmpty(shiftDirection))
            throw new ArgumentException("shiftDirection is required for delete_cells operation");

        var workbook = ctx.Document;
        var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);

        var rangeObj = ExcelHelper.CreateRange(worksheet.Cells, range);
        var shiftType = string.Equals(shiftDirection, "left", StringComparison.OrdinalIgnoreCase)
            ? ShiftType.Left
            : ShiftType.Up;

        worksheet.Cells.DeleteRange(
            rangeObj.FirstRow,
            rangeObj.FirstColumn,
            rangeObj.FirstRow + rangeObj.RowCount - 1,
            rangeObj.FirstColumn + rangeObj.ColumnCount - 1,
            shiftType);

        ctx.Save(outputPath);
        return $"Cells deleted in range {range}, shifted {shiftDirection}. {ctx.GetOutputMessage(outputPath)}";
    }
}