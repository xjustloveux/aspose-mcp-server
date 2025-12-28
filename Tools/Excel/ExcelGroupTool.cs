using System.Text.Json.Nodes;
using Aspose.Cells;
using AsposeMcpServer.Core;

namespace AsposeMcpServer.Tools.Excel;

/// <summary>
///     Unified tool for managing Excel row and column groups (group/ungroup).
/// </summary>
public class ExcelGroupTool : IAsposeTool
{
    /// <summary>
    ///     Gets the description of the tool and its usage examples.
    /// </summary>
    public string Description =>
        @"Manage Excel groups. Supports 4 operations: group_rows, ungroup_rows, group_columns, ungroup_columns.

Usage examples:
- Group rows: excel_group(operation='group_rows', path='book.xlsx', startRow=1, endRow=5)
- Ungroup rows: excel_group(operation='ungroup_rows', path='book.xlsx', startRow=1, endRow=5)
- Group columns: excel_group(operation='group_columns', path='book.xlsx', startColumn=1, endColumn=3)
- Ungroup columns: excel_group(operation='ungroup_columns', path='book.xlsx', startColumn=1, endColumn=3)";

    /// <summary>
    ///     Gets the JSON schema defining the input parameters for the tool.
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
                description = "End row index (0-based, must be >= startRow, required for group_rows/ungroup_rows)"
            },
            startColumn = new
            {
                type = "number",
                description = "Start column index (0-based, required for group_columns/ungroup_columns)"
            },
            endColumn = new
            {
                type = "number",
                description =
                    "End column index (0-based, must be >= startColumn, required for group_columns/ungroup_columns)"
            },
            isCollapsed = new
            {
                type = "boolean",
                description = "Collapse group initially (optional, for group_rows/group_columns, default: false)"
            },
            outputPath = new
            {
                type = "string",
                description = "Output file path (optional, defaults to input path)"
            }
        },
        required = new[] { "operation", "path" }
    };

    /// <summary>
    ///     Executes the tool operation with the provided JSON arguments.
    /// </summary>
    /// <param name="arguments">JSON arguments object containing operation parameters.</param>
    /// <returns>Result message as a string.</returns>
    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var operation = ArgumentHelper.GetString(arguments, "operation");
        var path = ArgumentHelper.GetAndValidatePath(arguments);
        var outputPath = ArgumentHelper.GetAndValidateOutputPath(arguments, path);
        var sheetIndex = ArgumentHelper.GetInt(arguments, "sheetIndex", 0);

        return operation.ToLower() switch
        {
            "group_rows" => await GroupRowsAsync(path, outputPath, sheetIndex, arguments),
            "ungroup_rows" => await UngroupRowsAsync(path, outputPath, sheetIndex, arguments),
            "group_columns" => await GroupColumnsAsync(path, outputPath, sheetIndex, arguments),
            "ungroup_columns" => await UngroupColumnsAsync(path, outputPath, sheetIndex, arguments),
            _ => throw new ArgumentException($"Unknown operation: {operation}")
        };
    }

    /// <summary>
    ///     Groups rows together.
    /// </summary>
    /// <param name="path">Excel file path.</param>
    /// <param name="outputPath">Output file path.</param>
    /// <param name="sheetIndex">Worksheet index (0-based).</param>
    /// <param name="arguments">JSON arguments containing startRow, endRow, optional isCollapsed.</param>
    /// <returns>Success message.</returns>
    private Task<string> GroupRowsAsync(string path, string outputPath, int sheetIndex, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var startRow = GetRequiredInt(arguments, "startRow", "group_rows");
            var endRow = GetRequiredInt(arguments, "endRow", "group_rows");
            var isCollapsed = ArgumentHelper.GetBool(arguments, "isCollapsed", false);

            ValidateRowRange(startRow, endRow);

            using var workbook = new Workbook(path);
            var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);
            worksheet.Cells.GroupRows(startRow, endRow, isCollapsed);

            workbook.Save(outputPath);
            return $"Rows {startRow}-{endRow} grouped in sheet {sheetIndex}. Output: {outputPath}";
        });
    }

    /// <summary>
    ///     Ungroups rows.
    /// </summary>
    /// <param name="path">Excel file path.</param>
    /// <param name="outputPath">Output file path.</param>
    /// <param name="sheetIndex">Worksheet index (0-based).</param>
    /// <param name="arguments">JSON arguments containing startRow, endRow.</param>
    /// <returns>Success message.</returns>
    private Task<string> UngroupRowsAsync(string path, string outputPath, int sheetIndex, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var startRow = GetRequiredInt(arguments, "startRow", "ungroup_rows");
            var endRow = GetRequiredInt(arguments, "endRow", "ungroup_rows");

            ValidateRowRange(startRow, endRow);

            using var workbook = new Workbook(path);
            var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);
            worksheet.Cells.UngroupRows(startRow, endRow);

            workbook.Save(outputPath);
            return $"Rows {startRow}-{endRow} ungrouped in sheet {sheetIndex}. Output: {outputPath}";
        });
    }

    /// <summary>
    ///     Groups columns together.
    /// </summary>
    /// <param name="path">Excel file path.</param>
    /// <param name="outputPath">Output file path.</param>
    /// <param name="sheetIndex">Worksheet index (0-based).</param>
    /// <param name="arguments">JSON arguments containing startColumn, endColumn, optional isCollapsed.</param>
    /// <returns>Success message.</returns>
    private Task<string> GroupColumnsAsync(string path, string outputPath, int sheetIndex, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var startColumn = GetRequiredInt(arguments, "startColumn", "group_columns");
            var endColumn = GetRequiredInt(arguments, "endColumn", "group_columns");
            var isCollapsed = ArgumentHelper.GetBool(arguments, "isCollapsed", false);

            ValidateColumnRange(startColumn, endColumn);

            using var workbook = new Workbook(path);
            var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);
            worksheet.Cells.GroupColumns(startColumn, endColumn, isCollapsed);

            workbook.Save(outputPath);
            return $"Columns {startColumn}-{endColumn} grouped in sheet {sheetIndex}. Output: {outputPath}";
        });
    }

    /// <summary>
    ///     Ungroups columns.
    /// </summary>
    /// <param name="path">Excel file path.</param>
    /// <param name="outputPath">Output file path.</param>
    /// <param name="sheetIndex">Worksheet index (0-based).</param>
    /// <param name="arguments">JSON arguments containing startColumn, endColumn.</param>
    /// <returns>Success message.</returns>
    private Task<string> UngroupColumnsAsync(string path, string outputPath, int sheetIndex, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var startColumn = GetRequiredInt(arguments, "startColumn", "ungroup_columns");
            var endColumn = GetRequiredInt(arguments, "endColumn", "ungroup_columns");

            ValidateColumnRange(startColumn, endColumn);

            using var workbook = new Workbook(path);
            var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);
            worksheet.Cells.UngroupColumns(startColumn, endColumn);

            workbook.Save(outputPath);
            return $"Columns {startColumn}-{endColumn} ungrouped in sheet {sheetIndex}. Output: {outputPath}";
        });
    }

    /// <summary>
    ///     Gets a required integer parameter and throws if not provided.
    /// </summary>
    /// <param name="arguments">JSON arguments.</param>
    /// <param name="paramName">Parameter name.</param>
    /// <param name="operationName">Operation name for error message.</param>
    /// <returns>Parameter value.</returns>
    /// <exception cref="ArgumentException">Thrown if parameter is not provided.</exception>
    private static int GetRequiredInt(JsonObject? arguments, string paramName, string operationName)
    {
        var value = ArgumentHelper.GetIntNullable(arguments, paramName);
        if (!value.HasValue)
            throw new ArgumentException($"Operation '{operationName}' requires parameter '{paramName}'.");
        return value.Value;
    }

    /// <summary>
    ///     Validates row range indices.
    /// </summary>
    /// <param name="startRow">Start row index.</param>
    /// <param name="endRow">End row index.</param>
    /// <exception cref="ArgumentException">Thrown if indices are invalid.</exception>
    private static void ValidateRowRange(int startRow, int endRow)
    {
        if (startRow < 0)
            throw new ArgumentException($"startRow cannot be negative. Got: {startRow}");
        if (endRow < 0)
            throw new ArgumentException($"endRow cannot be negative. Got: {endRow}");
        if (startRow > endRow)
            throw new ArgumentException($"startRow ({startRow}) cannot be greater than endRow ({endRow}).");
    }

    /// <summary>
    ///     Validates column range indices.
    /// </summary>
    /// <param name="startColumn">Start column index.</param>
    /// <param name="endColumn">End column index.</param>
    /// <exception cref="ArgumentException">Thrown if indices are invalid.</exception>
    private static void ValidateColumnRange(int startColumn, int endColumn)
    {
        if (startColumn < 0)
            throw new ArgumentException($"startColumn cannot be negative. Got: {startColumn}");
        if (endColumn < 0)
            throw new ArgumentException($"endColumn cannot be negative. Got: {endColumn}");
        if (startColumn > endColumn)
            throw new ArgumentException($"startColumn ({startColumn}) cannot be greater than endColumn ({endColumn}).");
    }
}