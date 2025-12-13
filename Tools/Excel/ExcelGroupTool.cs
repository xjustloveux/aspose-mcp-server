using System.Text.Json.Nodes;
using Aspose.Cells;
using AsposeMcpServer.Core;

namespace AsposeMcpServer.Tools;

/// <summary>
/// Unified tool for managing Excel groups (group/ungroup rows and columns)
/// Merges: ExcelGroupRowsTool, ExcelUngroupRowsTool, ExcelGroupColumnsTool, ExcelUngroupColumnsTool
/// </summary>
public class ExcelGroupTool : IAsposeTool
{
    public string Description => @"Manage Excel groups. Supports 4 operations: group_rows, ungroup_rows, group_columns, ungroup_columns.

Usage examples:
- Group rows: excel_group(operation='group_rows', path='book.xlsx', startRow=1, endRow=5)
- Ungroup rows: excel_group(operation='ungroup_rows', path='book.xlsx', startRow=1, endRow=5)
- Group columns: excel_group(operation='group_columns', path='book.xlsx', startColumn=1, endColumn=3)
- Ungroup columns: excel_group(operation='ungroup_columns', path='book.xlsx', startColumn=1, endColumn=3)";

    public object InputSchema => new
    {
        type = "object",
        properties = new
        {
            operation = new
            {
                type = "string",
                description = @"Operation to perform.
- 'group_rows': Group rows (required params: path, startRow, endRow)
- 'ungroup_rows': Ungroup rows (required params: path, startRow, endRow)
- 'group_columns': Group columns (required params: path, startColumn, endColumn)
- 'ungroup_columns': Ungroup columns (required params: path, startColumn, endColumn)",
                @enum = new[] { "group_rows", "ungroup_rows", "group_columns", "ungroup_columns" }
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
            startRow = new
            {
                type = "number",
                description = "Start row index (0-based, required for group_rows/ungroup_rows)"
            },
            endRow = new
            {
                type = "number",
                description = "End row index (0-based, required for group_rows/ungroup_rows)"
            },
            startColumn = new
            {
                type = "number",
                description = "Start column index (0-based, required for group_columns/ungroup_columns)"
            },
            endColumn = new
            {
                type = "number",
                description = "End column index (0-based, required for group_columns/ungroup_columns)"
            },
            isCollapsed = new
            {
                type = "boolean",
                description = "Collapse group initially (optional, for group_rows/group_columns, default: false)"
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
            "group_rows" => await GroupRowsAsync(arguments, path, sheetIndex),
            "ungroup_rows" => await UngroupRowsAsync(arguments, path, sheetIndex),
            "group_columns" => await GroupColumnsAsync(arguments, path, sheetIndex),
            "ungroup_columns" => await UngroupColumnsAsync(arguments, path, sheetIndex),
            _ => throw new ArgumentException($"Unknown operation: {operation}")
        };
    }

    private async Task<string> GroupRowsAsync(JsonObject? arguments, string path, int sheetIndex)
    {
        var startRow = arguments?["startRow"]?.GetValue<int>() ?? throw new ArgumentException("startRow is required for group_rows operation");
        var endRow = arguments?["endRow"]?.GetValue<int>() ?? throw new ArgumentException("endRow is required for group_rows operation");
        var isCollapsed = arguments?["isCollapsed"]?.GetValue<bool?>() ?? false;

        using var workbook = new Workbook(path);
        var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);
        worksheet.Cells.GroupRows(startRow, endRow, isCollapsed);

        workbook.Save(path);
        return await Task.FromResult($"Rows {startRow}-{endRow} grouped in sheet {sheetIndex}: {path}");
    }

    private async Task<string> UngroupRowsAsync(JsonObject? arguments, string path, int sheetIndex)
    {
        var startRow = arguments?["startRow"]?.GetValue<int>() ?? throw new ArgumentException("startRow is required for ungroup_rows operation");
        var endRow = arguments?["endRow"]?.GetValue<int>() ?? throw new ArgumentException("endRow is required for ungroup_rows operation");

        using var workbook = new Workbook(path);
        var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);
        worksheet.Cells.UngroupRows(startRow, endRow);

        workbook.Save(path);
        return await Task.FromResult($"Rows {startRow}-{endRow} ungrouped in sheet {sheetIndex}: {path}");
    }

    private async Task<string> GroupColumnsAsync(JsonObject? arguments, string path, int sheetIndex)
    {
        var startColumn = arguments?["startColumn"]?.GetValue<int>() ?? throw new ArgumentException("startColumn is required for group_columns operation");
        var endColumn = arguments?["endColumn"]?.GetValue<int>() ?? throw new ArgumentException("endColumn is required for group_columns operation");
        var isCollapsed = arguments?["isCollapsed"]?.GetValue<bool?>() ?? false;

        using var workbook = new Workbook(path);
        var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);
        worksheet.Cells.GroupColumns(startColumn, endColumn, isCollapsed);

        workbook.Save(path);
        return await Task.FromResult($"Columns {startColumn}-{endColumn} grouped in sheet {sheetIndex}: {path}");
    }

    private async Task<string> UngroupColumnsAsync(JsonObject? arguments, string path, int sheetIndex)
    {
        var startColumn = arguments?["startColumn"]?.GetValue<int>() ?? throw new ArgumentException("startColumn is required for ungroup_columns operation");
        var endColumn = arguments?["endColumn"]?.GetValue<int>() ?? throw new ArgumentException("endColumn is required for ungroup_columns operation");

        using var workbook = new Workbook(path);
        var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);
        worksheet.Cells.UngroupColumns(startColumn, endColumn);

        workbook.Save(path);
        return await Task.FromResult($"Columns {startColumn}-{endColumn} ungrouped in sheet {sheetIndex}: {path}");
    }
}

