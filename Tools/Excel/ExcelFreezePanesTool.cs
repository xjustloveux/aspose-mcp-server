using System.Text.Json;
using System.Text.Json.Nodes;
using Aspose.Cells;
using AsposeMcpServer.Core;

namespace AsposeMcpServer.Tools.Excel;

/// <summary>
///     Unified tool for managing Excel freeze panes (freeze/unfreeze/get).
/// </summary>
public class ExcelFreezePanesTool : IAsposeTool
{
    /// <summary>
    ///     Gets the description of the tool and its usage examples.
    /// </summary>
    public string Description => @"Manage Excel freeze panes. Supports 3 operations: freeze, unfreeze, get.

Usage examples:
- Freeze panes: excel_freeze_panes(operation='freeze', path='book.xlsx', row=1, column=1)
- Unfreeze panes: excel_freeze_panes(operation='unfreeze', path='book.xlsx')
- Get freeze status: excel_freeze_panes(operation='get', path='book.xlsx')";

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
- 'freeze': Freeze panes at specified row and column (required params: path, row, column)
- 'unfreeze': Remove freeze panes (required params: path)
- 'get': Get current freeze panes status (required params: path)",
                @enum = new[] { "freeze", "unfreeze", "get" }
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
            row = new
            {
                type = "number",
                description =
                    "Number of rows to freeze from top (0-based, required for freeze). E.g., row=2 freezes the first 2 rows."
            },
            column = new
            {
                type = "number",
                description =
                    "Number of columns to freeze from left (0-based, required for freeze). E.g., column=1 freezes the first column."
            },
            outputPath = new
            {
                type = "string",
                description = "Output file path (optional, for freeze/unfreeze operations, defaults to input path)"
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
            "freeze" => await FreezePanesAsync(path, outputPath, sheetIndex, arguments),
            "unfreeze" => await UnfreezePanesAsync(path, outputPath, sheetIndex),
            "get" => await GetFreezePanesAsync(path, sheetIndex),
            _ => throw new ArgumentException($"Unknown operation: {operation}")
        };
    }

    /// <summary>
    ///     Freezes panes at the specified row and column.
    /// </summary>
    /// <param name="path">Excel file path.</param>
    /// <param name="outputPath">Output file path.</param>
    /// <param name="sheetIndex">Worksheet index (0-based).</param>
    /// <param name="arguments">JSON arguments containing row and column (0-based).</param>
    /// <returns>Success message with freeze position.</returns>
    private Task<string> FreezePanesAsync(string path, string outputPath, int sheetIndex, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var row = ArgumentHelper.GetInt(arguments, "row");
            var column = ArgumentHelper.GetInt(arguments, "column");

            using var workbook = new Workbook(path);
            var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);

            // FreezePanes(row, column, frozenRows, frozenColumns)
            // - row/column: 1-based position where the freeze starts
            // - frozenRows/frozenColumns: number of frozen rows/columns visible in the frozen pane
            worksheet.FreezePanes(row + 1, column + 1, row, column);

            workbook.Save(outputPath);
            return $"Frozen panes at row {row}, column {column}. Output: {outputPath}";
        });
    }

    /// <summary>
    ///     Unfreezes panes in the worksheet.
    /// </summary>
    /// <param name="path">Excel file path.</param>
    /// <param name="outputPath">Output file path.</param>
    /// <param name="sheetIndex">Worksheet index (0-based).</param>
    /// <returns>Success message.</returns>
    private Task<string> UnfreezePanesAsync(string path, string outputPath, int sheetIndex)
    {
        return Task.Run(() =>
        {
            using var workbook = new Workbook(path);
            var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);

            worksheet.UnFreezePanes();

            workbook.Save(outputPath);
            return $"Unfrozen panes. Output: {outputPath}";
        });
    }

    /// <summary>
    ///     Gets freeze panes status for the worksheet.
    /// </summary>
    /// <param name="path">Excel file path.</param>
    /// <param name="sheetIndex">Worksheet index (0-based).</param>
    /// <returns>JSON string with freeze panes status.</returns>
    private Task<string> GetFreezePanesAsync(string path, int sheetIndex)
    {
        return Task.Run(() =>
        {
            using var workbook = new Workbook(path);
            var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);

            var isFrozen = worksheet.PaneState == PaneStateType.Frozen;
            int? frozenRow = null;
            int? frozenColumn = null;
            int? frozenRows = null;
            int? frozenColumns = null;

            if (isFrozen)
            {
                worksheet.GetFreezedPanes(out var row, out var col, out var rows, out var cols);
                frozenRow = row > 0 ? row - 1 : 0;
                frozenColumn = col > 0 ? col - 1 : 0;
                frozenRows = rows;
                frozenColumns = cols;
            }

            var result = new
            {
                worksheetName = worksheet.Name,
                isFrozen,
                frozenRow,
                frozenColumn,
                frozenRows,
                frozenColumns,
                status = isFrozen ? "Panes are frozen" : "Panes are not frozen"
            };

            return JsonSerializer.Serialize(result, new JsonSerializerOptions { WriteIndented = true });
        });
    }
}