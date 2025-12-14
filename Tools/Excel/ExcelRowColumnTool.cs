using System.Text.Json.Nodes;
using Aspose.Cells;
using AsposeMcpServer.Core;

namespace AsposeMcpServer.Tools;

/// <summary>
/// Unified tool for managing Excel rows and columns (insert/delete rows, columns, cells)
/// Merges: ExcelInsertRowTool, ExcelDeleteRowTool, ExcelInsertColumnTool, ExcelDeleteColumnTool, 
/// ExcelInsertCellsTool, ExcelDeleteCellsTool
/// </summary>
public class ExcelRowColumnTool : IAsposeTool
{
    public string Description => @"Manage Excel rows and columns. Supports 6 operations: insert_row, delete_row, insert_column, delete_column, insert_cells, delete_cells.

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
                @enum = new[] { "insert_row", "delete_row", "insert_column", "delete_column", "insert_cells", "delete_cells" }
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
            }
        },
        required = new[] { "operation", "path" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var operation = ArgumentHelper.GetString(arguments, "operation", "operation");
        var path = ArgumentHelper.GetAndValidatePath(arguments);
        var sheetIndex = arguments?["sheetIndex"]?.GetValue<int>() ?? 0;

        return operation.ToLower() switch
        {
            "insert_row" => await InsertRowAsync(arguments, path, sheetIndex),
            "delete_row" => await DeleteRowAsync(arguments, path, sheetIndex),
            "insert_column" => await InsertColumnAsync(arguments, path, sheetIndex),
            "delete_column" => await DeleteColumnAsync(arguments, path, sheetIndex),
            "insert_cells" => await InsertCellsAsync(arguments, path, sheetIndex),
            "delete_cells" => await DeleteCellsAsync(arguments, path, sheetIndex),
            _ => throw new ArgumentException($"Unknown operation: {operation}")
        };
    }

    /// <summary>
    /// Inserts rows into the worksheet
    /// </summary>
    /// <param name="arguments">JSON arguments containing rowIndex, optional count</param>
    /// <param name="path">Excel file path</param>
    /// <param name="sheetIndex">Worksheet index (0-based)</param>
    /// <returns>Success message</returns>
    private async Task<string> InsertRowAsync(JsonObject? arguments, string path, int sheetIndex)
    {
        var rowIndex = ArgumentHelper.GetInt(arguments, "rowIndex", "rowIndex");
        var count = arguments?["count"]?.GetValue<int>() ?? 1;

        using var workbook = new Workbook(path);
        var worksheet = workbook.Worksheets[sheetIndex];

        for (int i = 0; i < count; i++)
        {
            worksheet.Cells.InsertRow(rowIndex);
        }
        workbook.Save(path);

        return await Task.FromResult($"在第 {rowIndex} 行插入了 {count} 行: {path}");
    }

    /// <summary>
    /// Deletes rows from the worksheet
    /// </summary>
    /// <param name="arguments">JSON arguments containing rowIndex, optional count</param>
    /// <param name="path">Excel file path</param>
    /// <param name="sheetIndex">Worksheet index (0-based)</param>
    /// <returns>Success message</returns>
    private async Task<string> DeleteRowAsync(JsonObject? arguments, string path, int sheetIndex)
    {
        var rowIndex = ArgumentHelper.GetInt(arguments, "rowIndex", "rowIndex");
        var count = arguments?["count"]?.GetValue<int>() ?? 1;

        using var workbook = new Workbook(path);
        var worksheet = workbook.Worksheets[sheetIndex];

        for (int i = 0; i < count; i++)
        {
            worksheet.Cells.DeleteRow(rowIndex);
        }
        workbook.Save(path);

        return await Task.FromResult($"已刪除第 {rowIndex} 行起的 {count} 行: {path}");
    }

    /// <summary>
    /// Inserts columns into the worksheet
    /// </summary>
    /// <param name="arguments">JSON arguments containing columnIndex, optional count</param>
    /// <param name="path">Excel file path</param>
    /// <param name="sheetIndex">Worksheet index (0-based)</param>
    /// <returns>Success message</returns>
    private async Task<string> InsertColumnAsync(JsonObject? arguments, string path, int sheetIndex)
    {
        var columnIndex = ArgumentHelper.GetInt(arguments, "columnIndex", "columnIndex");
        var count = arguments?["count"]?.GetValue<int>() ?? 1;

        using var workbook = new Workbook(path);
        var worksheet = workbook.Worksheets[sheetIndex];

        for (int i = 0; i < count; i++)
        {
            worksheet.Cells.InsertColumn(columnIndex);
        }
        workbook.Save(path);

        return await Task.FromResult($"在第 {columnIndex} 列插入了 {count} 列: {path}");
    }

    /// <summary>
    /// Deletes columns from the worksheet
    /// </summary>
    /// <param name="arguments">JSON arguments containing columnIndex, optional count</param>
    /// <param name="path">Excel file path</param>
    /// <param name="sheetIndex">Worksheet index (0-based)</param>
    /// <returns>Success message</returns>
    private async Task<string> DeleteColumnAsync(JsonObject? arguments, string path, int sheetIndex)
    {
        var columnIndex = ArgumentHelper.GetInt(arguments, "columnIndex", "columnIndex");
        var count = arguments?["count"]?.GetValue<int>() ?? 1;

        using var workbook = new Workbook(path);
        var worksheet = workbook.Worksheets[sheetIndex];

        for (int i = 0; i < count; i++)
        {
            worksheet.Cells.DeleteColumn(columnIndex);
        }
        workbook.Save(path);

        return await Task.FromResult($"已刪除第 {columnIndex} 列起的 {count} 列: {path}");
    }

    /// <summary>
    /// Inserts cells into the worksheet
    /// </summary>
    /// <param name="arguments">JSON arguments containing range, shiftDirection</param>
    /// <param name="path">Excel file path</param>
    /// <param name="sheetIndex">Worksheet index (0-based)</param>
    /// <returns>Success message</returns>
    private async Task<string> InsertCellsAsync(JsonObject? arguments, string path, int sheetIndex)
    {
        var range = ArgumentHelper.GetString(arguments, "range", "range");
        var shiftDirection = ArgumentHelper.GetString(arguments, "shiftDirection", "shiftDirection");

        using var workbook = new Workbook(path);
        var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);
        var rangeObj = worksheet.Cells.CreateRange(range);

        var shiftType = shiftDirection.ToLower() == "right" ? ShiftType.Right : ShiftType.Down;
        
        if (shiftType == ShiftType.Down)
        {
            for (int i = 0; i < rangeObj.RowCount; i++)
            {
                worksheet.Cells.InsertRow(rangeObj.FirstRow);
            }
        }
        else
        {
            for (int i = 0; i < rangeObj.ColumnCount; i++)
            {
                worksheet.Cells.InsertColumn(rangeObj.FirstColumn);
            }
        }

        workbook.Save(path);
        return await Task.FromResult($"Cells inserted in range {range}, shifted {shiftDirection}: {path}");
    }

    /// <summary>
    /// Deletes cells from the worksheet
    /// </summary>
    /// <param name="arguments">JSON arguments containing range, shiftDirection</param>
    /// <param name="path">Excel file path</param>
    /// <param name="sheetIndex">Worksheet index (0-based)</param>
    /// <returns>Success message</returns>
    private async Task<string> DeleteCellsAsync(JsonObject? arguments, string path, int sheetIndex)
    {
        var range = ArgumentHelper.GetString(arguments, "range", "range");
        var shiftDirection = ArgumentHelper.GetString(arguments, "shiftDirection", "shiftDirection");

        using var workbook = new Workbook(path);
        var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);
        var rangeObj = worksheet.Cells.CreateRange(range);

        var shiftType = shiftDirection.ToLower() == "left" ? ShiftType.Left : ShiftType.Up;
        worksheet.Cells.DeleteRange(rangeObj.FirstRow, rangeObj.FirstColumn, rangeObj.RowCount, rangeObj.ColumnCount, shiftType);

        workbook.Save(path);
        return await Task.FromResult($"Cells deleted in range {range}, shifted {shiftDirection}: {path}");
    }
}

