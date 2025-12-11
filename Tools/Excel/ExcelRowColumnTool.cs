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
    public string Description => "Manage Excel rows and columns: insert or delete rows, columns, or cells";

    public object InputSchema => new
    {
        type = "object",
        properties = new
        {
            operation = new
            {
                type = "string",
                description = "Operation to perform: 'insert_row', 'delete_row', 'insert_column', 'delete_column', 'insert_cells', 'delete_cells'",
                @enum = new[] { "insert_row", "delete_row", "insert_column", "delete_column", "insert_cells", "delete_cells" }
            },
            path = new
            {
                type = "string",
                description = "Excel file path"
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
        var operation = arguments?["operation"]?.GetValue<string>() ?? throw new ArgumentException("operation is required");
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        SecurityHelper.ValidateFilePath(path, "path");
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

    private async Task<string> InsertRowAsync(JsonObject? arguments, string path, int sheetIndex)
    {
        var rowIndex = arguments?["rowIndex"]?.GetValue<int>() ?? throw new ArgumentException("rowIndex is required for insert_row operation");
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

    private async Task<string> DeleteRowAsync(JsonObject? arguments, string path, int sheetIndex)
    {
        var rowIndex = arguments?["rowIndex"]?.GetValue<int>() ?? throw new ArgumentException("rowIndex is required for delete_row operation");
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

    private async Task<string> InsertColumnAsync(JsonObject? arguments, string path, int sheetIndex)
    {
        var columnIndex = arguments?["columnIndex"]?.GetValue<int>() ?? throw new ArgumentException("columnIndex is required for insert_column operation");
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

    private async Task<string> DeleteColumnAsync(JsonObject? arguments, string path, int sheetIndex)
    {
        var columnIndex = arguments?["columnIndex"]?.GetValue<int>() ?? throw new ArgumentException("columnIndex is required for delete_column operation");
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

    private async Task<string> InsertCellsAsync(JsonObject? arguments, string path, int sheetIndex)
    {
        var range = arguments?["range"]?.GetValue<string>() ?? throw new ArgumentException("range is required for insert_cells operation");
        var shiftDirection = arguments?["shiftDirection"]?.GetValue<string>() ?? throw new ArgumentException("shiftDirection is required for insert_cells operation");

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

    private async Task<string> DeleteCellsAsync(JsonObject? arguments, string path, int sheetIndex)
    {
        var range = arguments?["range"]?.GetValue<string>() ?? throw new ArgumentException("range is required for delete_cells operation");
        var shiftDirection = arguments?["shiftDirection"]?.GetValue<string>() ?? throw new ArgumentException("shiftDirection is required for delete_cells operation");

        using var workbook = new Workbook(path);
        var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);
        var rangeObj = worksheet.Cells.CreateRange(range);

        var shiftType = shiftDirection.ToLower() == "left" ? ShiftType.Left : ShiftType.Up;
        worksheet.Cells.DeleteRange(rangeObj.FirstRow, rangeObj.FirstColumn, rangeObj.RowCount, rangeObj.ColumnCount, shiftType);

        workbook.Save(path);
        return await Task.FromResult($"Cells deleted in range {range}, shifted {shiftDirection}: {path}");
    }
}

