using System.Text.Json.Nodes;
using Aspose.Cells;
using AsposeMcpServer.Core;

namespace AsposeMcpServer.Tools.Excel;

/// <summary>
///     Unified tool for managing Excel groups (group/ungroup rows and columns)
///     Merges: ExcelGroupRowsTool, ExcelUngroupRowsTool, ExcelGroupColumnsTool, ExcelUngroupColumnsTool
/// </summary>
public class ExcelGroupTool : IAsposeTool
{
    /// <summary>
    ///     Gets the description of the tool and its usage examples
    /// </summary>
    public string Description =>
        @"Manage Excel groups. Supports 4 operations: group_rows, ungroup_rows, group_columns, ungroup_columns.

Usage examples:
- Group rows: excel_group(operation='group_rows', path='book.xlsx', startRow=1, endRow=5)
- Ungroup rows: excel_group(operation='ungroup_rows', path='book.xlsx', startRow=1, endRow=5)
- Group columns: excel_group(operation='group_columns', path='book.xlsx', startColumn=1, endColumn=3)
- Ungroup columns: excel_group(operation='ungroup_columns', path='book.xlsx', startColumn=1, endColumn=3)";

    /// <summary>
    ///     Gets the JSON schema defining the input parameters for the tool
    /// </summary>
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
        var sheetIndex = ArgumentHelper.GetInt(arguments, "sheetIndex", 0);

        return operation.ToLower() switch
        {
            "group_rows" => await GroupRowsAsync(arguments, path, sheetIndex),
            "ungroup_rows" => await UngroupRowsAsync(arguments, path, sheetIndex),
            "group_columns" => await GroupColumnsAsync(arguments, path, sheetIndex),
            "ungroup_columns" => await UngroupColumnsAsync(arguments, path, sheetIndex),
            _ => throw new ArgumentException($"Unknown operation: {operation}")
        };
    }

    /// <summary>
    ///     Groups rows together
    /// </summary>
    /// <param name="arguments">JSON arguments containing startRow, endRow, optional isCollapsed</param>
    /// <param name="path">Excel file path</param>
    /// <param name="sheetIndex">Worksheet index (0-based)</param>
    /// <returns>Success message</returns>
    private Task<string> GroupRowsAsync(JsonObject? arguments, string path, int sheetIndex)
    {
        return Task.Run(() =>
        {
            var outputPath = ArgumentHelper.GetAndValidateOutputPath(arguments, path);
            var startRow = ArgumentHelper.GetInt(arguments, "startRow");
            var endRow = ArgumentHelper.GetInt(arguments, "endRow");
            var isCollapsed = ArgumentHelper.GetBool(arguments, "isCollapsed", false);

            using var workbook = new Workbook(path);
            var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);
            worksheet.Cells.GroupRows(startRow, endRow, isCollapsed);

            workbook.Save(outputPath);
            return $"Rows {startRow}-{endRow} grouped in sheet {sheetIndex}: {outputPath}";
        });
    }

    /// <summary>
    ///     Ungroups rows
    /// </summary>
    /// <param name="arguments">JSON arguments containing startRow, endRow</param>
    /// <param name="path">Excel file path</param>
    /// <param name="sheetIndex">Worksheet index (0-based)</param>
    /// <returns>Success message</returns>
    private Task<string> UngroupRowsAsync(JsonObject? arguments, string path, int sheetIndex)
    {
        return Task.Run(() =>
        {
            var outputPath = ArgumentHelper.GetAndValidateOutputPath(arguments, path);
            var startRow = ArgumentHelper.GetInt(arguments, "startRow");
            var endRow = ArgumentHelper.GetInt(arguments, "endRow");

            using var workbook = new Workbook(path);
            var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);
            worksheet.Cells.UngroupRows(startRow, endRow);

            workbook.Save(outputPath);
            return $"Rows {startRow}-{endRow} ungrouped in sheet {sheetIndex}: {outputPath}";
        });
    }

    /// <summary>
    ///     Groups columns together
    /// </summary>
    /// <param name="arguments">JSON arguments containing startColumn, endColumn, optional isCollapsed</param>
    /// <param name="path">Excel file path</param>
    /// <param name="sheetIndex">Worksheet index (0-based)</param>
    /// <returns>Success message</returns>
    private Task<string> GroupColumnsAsync(JsonObject? arguments, string path, int sheetIndex)
    {
        return Task.Run(() =>
        {
            var outputPath = ArgumentHelper.GetAndValidateOutputPath(arguments, path);
            var startColumn = ArgumentHelper.GetInt(arguments, "startColumn");
            var endColumn = ArgumentHelper.GetInt(arguments, "endColumn");
            var isCollapsed = ArgumentHelper.GetBool(arguments, "isCollapsed", false);

            using var workbook = new Workbook(path);
            var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);
            worksheet.Cells.GroupColumns(startColumn, endColumn, isCollapsed);

            workbook.Save(outputPath);
            return $"Columns {startColumn}-{endColumn} grouped in sheet {sheetIndex}: {outputPath}";
        });
    }

    /// <summary>
    ///     Ungroups columns
    /// </summary>
    /// <param name="arguments">JSON arguments containing startColumn, endColumn</param>
    /// <param name="path">Excel file path</param>
    /// <param name="sheetIndex">Worksheet index (0-based)</param>
    /// <returns>Success message</returns>
    private Task<string> UngroupColumnsAsync(JsonObject? arguments, string path, int sheetIndex)
    {
        return Task.Run(() =>
        {
            var outputPath = ArgumentHelper.GetAndValidateOutputPath(arguments, path);
            var startColumn = ArgumentHelper.GetInt(arguments, "startColumn");
            var endColumn = ArgumentHelper.GetInt(arguments, "endColumn");

            using var workbook = new Workbook(path);
            var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);
            worksheet.Cells.UngroupColumns(startColumn, endColumn);

            workbook.Save(outputPath);
            return $"Columns {startColumn}-{endColumn} ungrouped in sheet {sheetIndex}: {outputPath}";
        });
    }
}