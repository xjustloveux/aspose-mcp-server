using System.Text.Json.Nodes;
using Aspose.Cells;
using AsposeMcpServer.Core;
using Range = Aspose.Cells.Range;

namespace AsposeMcpServer.Tools.Excel;

/// <summary>
///     Unified tool for managing Excel rows and columns (insert/delete rows, columns, cells)
///     Merges: ExcelInsertRowTool, ExcelDeleteRowTool, ExcelInsertColumnTool, ExcelDeleteColumnTool,
///     ExcelInsertCellsTool, ExcelDeleteCellsTool
/// </summary>
public class ExcelRowColumnTool : IAsposeTool
{
    public string Description =>
        @"Manage Excel rows and columns. Supports 6 operations: insert_row, delete_row, insert_column, delete_column, insert_cells, delete_cells.

Usage examples:
- Insert row: excel_row_column(operation='insert_row', path='book.xlsx', rowIndex=2, count=1)
- Delete row: excel_row_column(operation='delete_row', path='book.xlsx', rowIndex=2)
- Insert column: excel_row_column(operation='insert_column', path='book.xlsx', columnIndex=2, count=1)
- Delete column: excel_row_column(operation='delete_column', path='book.xlsx', columnIndex=2)
- Insert cells: excel_row_column(operation='insert_cells', path='book.xlsx', range='A1:C5', shiftDirection='Down')
- Delete cells: excel_row_column(operation='delete_cells', path='book.xlsx', range='A1:C5', shiftDirection='Up')";

    public object InputSchema => new
    {
        type = "object",
        properties = new
        {
            operation = new
            {
                type = "string",
                description = @"Operation to perform.
- 'insert_row': Insert row(s) (required params: path, rowIndex)
- 'delete_row': Delete row(s) (required params: path, rowIndex)
- 'insert_column': Insert column(s) (required params: path, columnIndex)
- 'delete_column': Delete column(s) (required params: path, columnIndex)
- 'insert_cells': Insert cells (required params: path, range, shiftDirection)
- 'delete_cells': Delete cells (required params: path, range, shiftDirection)",
                @enum = new[]
                    { "insert_row", "delete_row", "insert_column", "delete_column", "insert_cells", "delete_cells" }
            },
            path = new
            {
                type = "string",
                description = "Excel file path (required for all operations)"
            },
            sheetIndex = new
            {
                type = "number",
                description = "Sheet index (0-based, optional, default: 0)"
            },
            rowIndex = new
            {
                type = "number",
                description = "Row index (0-based, required for insert_row/delete_row)"
            },
            columnIndex = new
            {
                type = "number",
                description = "Column index (0-based, required for insert_column/delete_column)"
            },
            range = new
            {
                type = "string",
                description = "Cell range (e.g., 'A1:C5', required for insert_cells/delete_cells)"
            },
            count = new
            {
                type = "number",
                description = "Number of rows/columns to insert/delete (optional, default: 1)"
            },
            shiftDirection = new
            {
                type = "string",
                description = "Shift direction: 'Right'/'Down' for insert_cells, 'Left'/'Up' for delete_cells",
                @enum = new[] { "Right", "Down", "Left", "Up" }
            },
            outputPath = new
            {
                type = "string",
                description = "Output file path (optional, for all operations, defaults to input path)"
            }
        },
        required = new[] { "operation", "path" }
    };

    /// <summary>
    ///     Executes the tool operation with the provided JSON arguments
    /// </summary>
    /// <param name="arguments">JSON arguments object containing operation parameters</param>
    /// <returns>Result message as a string</returns>
    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var operation = ArgumentHelper.GetString(arguments, "operation");
        var path = ArgumentHelper.GetAndValidatePath(arguments);
        var outputPath = ArgumentHelper.GetAndValidateOutputPath(arguments, path);
        var sheetIndex = ArgumentHelper.GetInt(arguments, "sheetIndex", 0);

        return operation.ToLower() switch
        {
            "insert_row" => await InsertRowAsync(path, outputPath, sheetIndex, arguments),
            "delete_row" => await DeleteRowAsync(path, outputPath, sheetIndex, arguments),
            "insert_column" => await InsertColumnAsync(path, outputPath, sheetIndex, arguments),
            "delete_column" => await DeleteColumnAsync(path, outputPath, sheetIndex, arguments),
            "insert_cells" => await InsertCellsAsync(path, outputPath, sheetIndex, arguments),
            "delete_cells" => await DeleteCellsAsync(path, outputPath, sheetIndex, arguments),
            "set_column_width" => throw new ArgumentException(
                $"Operation 'set_column_width' is not supported by excel_row_column. Please use excel_view_settings operation instead. Example: excel_view_settings(operation='set_column_width', path='{path}', columnIndex=0, width=15)"),
            _ => throw new ArgumentException($"Unknown operation: {operation}")
        };
    }

    /// <summary>
    ///     Inserts rows into the worksheet
    /// </summary>
    /// <param name="path">Excel file path</param>
    /// <param name="outputPath">Output file path</param>
    /// <param name="sheetIndex">Worksheet index (0-based)</param>
    /// <param name="arguments">JSON arguments containing rowIndex, optional count</param>
    /// <returns>Success message</returns>
    private Task<string> InsertRowAsync(string path, string outputPath, int sheetIndex, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var rowIndex = ArgumentHelper.GetInt(arguments, "rowIndex");
            var count = ArgumentHelper.GetInt(arguments, "count", 1);

            using var workbook = new Workbook(path);
            var worksheet = workbook.Worksheets[sheetIndex];

            for (var i = 0; i < count; i++) worksheet.Cells.InsertRow(rowIndex);
            workbook.Save(outputPath);

            return $"Inserted {count} row(s) at row {rowIndex}. Output: {outputPath}";
        });
    }

    /// <summary>
    ///     Deletes rows from the worksheet
    /// </summary>
    /// <param name="path">Excel file path</param>
    /// <param name="outputPath">Output file path</param>
    /// <param name="sheetIndex">Worksheet index (0-based)</param>
    /// <param name="arguments">JSON arguments containing rowIndex, optional count</param>
    /// <returns>Success message</returns>
    private Task<string> DeleteRowAsync(string path, string outputPath, int sheetIndex, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var rowIndex = ArgumentHelper.GetInt(arguments, "rowIndex");
            var count = ArgumentHelper.GetInt(arguments, "count", 1);

            using var workbook = new Workbook(path);
            var worksheet = workbook.Worksheets[sheetIndex];

            for (var i = 0; i < count; i++) worksheet.Cells.DeleteRow(rowIndex);
            workbook.Save(outputPath);

            return $"Deleted {count} row(s) starting from row {rowIndex}. Output: {outputPath}";
        });
    }

    /// <summary>
    ///     Inserts columns into the worksheet
    /// </summary>
    /// <param name="path">Excel file path</param>
    /// <param name="outputPath">Output file path</param>
    /// <param name="sheetIndex">Worksheet index (0-based)</param>
    /// <param name="arguments">JSON arguments containing columnIndex, optional count</param>
    /// <returns>Success message</returns>
    private Task<string> InsertColumnAsync(string path, string outputPath, int sheetIndex, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var columnIndex = ArgumentHelper.GetInt(arguments, "columnIndex");
            var count = ArgumentHelper.GetInt(arguments, "count", 1);

            using var workbook = new Workbook(path);
            var worksheet = workbook.Worksheets[sheetIndex];

            for (var i = 0; i < count; i++) worksheet.Cells.InsertColumn(columnIndex);
            workbook.Save(outputPath);

            return $"Inserted {count} column(s) at column {columnIndex}. Output: {outputPath}";
        });
    }

    /// <summary>
    ///     Deletes columns from the worksheet
    /// </summary>
    /// <param name="path">Excel file path</param>
    /// <param name="outputPath">Output file path</param>
    /// <param name="sheetIndex">Worksheet index (0-based)</param>
    /// <param name="arguments">JSON arguments containing columnIndex, optional count</param>
    /// <returns>Success message</returns>
    private Task<string> DeleteColumnAsync(string path, string outputPath, int sheetIndex, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var columnIndex = ArgumentHelper.GetInt(arguments, "columnIndex");
            var count = ArgumentHelper.GetInt(arguments, "count", 1);

            using var workbook = new Workbook(path);
            var worksheet = workbook.Worksheets[sheetIndex];

            for (var i = 0; i < count; i++) worksheet.Cells.DeleteColumn(columnIndex);
            workbook.Save(outputPath);

            return $"Deleted {count} column(s) starting from column {columnIndex}. Output: {outputPath}";
        });
    }

    /// <summary>
    ///     Inserts cells into the worksheet
    /// </summary>
    /// <param name="path">Excel file path</param>
    /// <param name="outputPath">Output file path</param>
    /// <param name="sheetIndex">Worksheet index (0-based)</param>
    /// <param name="arguments">JSON arguments containing range, shiftDirection</param>
    /// <returns>Success message</returns>
    private Task<string> InsertCellsAsync(string path, string outputPath, int sheetIndex, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var range = ArgumentHelper.GetString(arguments, "range");
            var shiftDirection = ArgumentHelper.GetString(arguments, "shiftDirection");

            using var workbook = new Workbook(path);
            var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);

            var rangeObj = ExcelHelper.CreateRange(worksheet.Cells, range);

            var shiftType = shiftDirection.ToLower() == "right" ? ShiftType.Right : ShiftType.Down;

            if (shiftType == ShiftType.Down)
                for (var i = 0; i < rangeObj.RowCount; i++)
                    worksheet.Cells.InsertRow(rangeObj.FirstRow);
            else
                for (var i = 0; i < rangeObj.ColumnCount; i++)
                    worksheet.Cells.InsertColumn(rangeObj.FirstColumn);

            workbook.Save(outputPath);
            return $"Cells inserted in range {range}, shifted {shiftDirection}. Output: {outputPath}";
        });
    }

    /// <summary>
    ///     Deletes cells from the worksheet
    /// </summary>
    /// <param name="path">Excel file path</param>
    /// <param name="outputPath">Output file path</param>
    /// <param name="sheetIndex">Worksheet index (0-based)</param>
    /// <param name="arguments">JSON arguments containing range, shiftDirection</param>
    /// <returns>Success message</returns>
    private Task<string> DeleteCellsAsync(string path, string outputPath, int sheetIndex, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var range = ArgumentHelper.GetString(arguments, "range");
            var shiftDirection = ArgumentHelper.GetString(arguments, "shiftDirection");

            using var workbook = new Workbook(path);
            var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);

            // Convert single cell to range format (e.g., "B27" -> "B27:B27") for proper deletion handling
            var normalizedRange = range;
            if (!range.Contains(':')) normalizedRange = $"{range}:{range}";

            Range rangeObj;
            try
            {
                rangeObj = ExcelHelper.CreateRange(worksheet.Cells, normalizedRange);
            }
            catch (Exception ex)
            {
                throw new ArgumentException(
                    $"Invalid range format: '{range}'. Single cell format (e.g., 'B2') or range format (e.g., 'B3:B3') is expected. Error: {ex.Message}");
            }

            // Ensure RowCount and ColumnCount are at least 1 for DeleteRange
            var rowCount = Math.Max(1, rangeObj.RowCount);
            var columnCount = Math.Max(1, rangeObj.ColumnCount);

            var shiftType = shiftDirection.ToLower() == "left" ? ShiftType.Left : ShiftType.Up;
            worksheet.Cells.DeleteRange(rangeObj.FirstRow, rangeObj.FirstColumn, rowCount, columnCount, shiftType);

            workbook.Save(outputPath);
            return $"Cells deleted in range {range}, shifted {shiftDirection}. Output: {outputPath}";
        });
    }
}