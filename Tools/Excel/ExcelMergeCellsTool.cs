using System.Text.Json;
using System.Text.Json.Nodes;
using Aspose.Cells;
using AsposeMcpServer.Core;

namespace AsposeMcpServer.Tools.Excel;

/// <summary>
///     Unified tool for managing Excel merged cells (merge, unmerge, get).
/// </summary>
public class ExcelMergeCellsTool : IAsposeTool
{
    /// <summary>
    ///     Gets the description of the tool and its usage examples.
    /// </summary>
    public string Description => @"Manage Excel merged cells. Supports 3 operations: merge, unmerge, get.

Usage examples:
- Merge cells: excel_merge_cells(operation='merge', path='book.xlsx', range='A1:C1')
- Unmerge cells: excel_merge_cells(operation='unmerge', path='book.xlsx', range='A1:C1')
- Get merged cells: excel_merge_cells(operation='get', path='book.xlsx')

WARNING: Merging cells will only keep the value of the top-left cell. All other cell values will be lost.";

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
- 'merge': Merge cells (required params: path, range). Note: only top-left cell value is kept.
- 'unmerge': Unmerge cells (required params: path, range)
- 'get': Get merged cells info (required params: path)",
                @enum = new[] { "merge", "unmerge", "get" }
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
            range = new
            {
                type = "string",
                description =
                    "Cell range to merge/unmerge (e.g., 'A1:C3', must include at least 2 cells, required for merge/unmerge)"
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
    /// <exception cref="ArgumentException">Thrown when operation is unknown.</exception>
    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var operation = ArgumentHelper.GetString(arguments, "operation");
        var path = ArgumentHelper.GetAndValidatePath(arguments);
        var outputPath = ArgumentHelper.GetAndValidateOutputPath(arguments, path);
        var sheetIndex = ArgumentHelper.GetInt(arguments, "sheetIndex", 0);

        return operation.ToLower() switch
        {
            "merge" => await MergeCellsAsync(path, outputPath, sheetIndex, arguments),
            "unmerge" => await UnmergeCellsAsync(path, outputPath, sheetIndex, arguments),
            "get" => await GetMergedCellsAsync(path, sheetIndex),
            _ => throw new ArgumentException($"Unknown operation: {operation}")
        };
    }

    /// <summary>
    ///     Merges cells in a range.
    /// </summary>
    /// <param name="path">Excel file path.</param>
    /// <param name="outputPath">Output file path.</param>
    /// <param name="sheetIndex">Worksheet index (0-based).</param>
    /// <param name="arguments">JSON arguments containing range.</param>
    /// <returns>Success message with merged range details.</returns>
    /// <exception cref="ArgumentException">Thrown when range contains only a single cell.</exception>
    private Task<string> MergeCellsAsync(string path, string outputPath, int sheetIndex, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var range = ArgumentHelper.GetString(arguments, "range");

            using var workbook = new Workbook(path);
            var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);
            var cells = worksheet.Cells;

            var cellRange = ExcelHelper.CreateRange(cells, range);

            if (cellRange is { RowCount: 1, ColumnCount: 1 })
                throw new ArgumentException(
                    $"Cannot merge a single cell. Range '{range}' must include at least 2 cells.");

            cellRange.Merge();
            workbook.Save(outputPath);

            return
                $"Range {range} merged ({cellRange.RowCount} rows x {cellRange.ColumnCount} columns). Output: {outputPath}";
        });
    }

    /// <summary>
    ///     Unmerges cells in a range.
    /// </summary>
    /// <param name="path">Excel file path.</param>
    /// <param name="outputPath">Output file path.</param>
    /// <param name="sheetIndex">Worksheet index (0-based).</param>
    /// <param name="arguments">JSON arguments containing range.</param>
    /// <returns>Success message.</returns>
    private Task<string> UnmergeCellsAsync(string path, string outputPath, int sheetIndex, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var range = ArgumentHelper.GetString(arguments, "range");

            using var workbook = new Workbook(path);
            var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);
            var cells = worksheet.Cells;

            var cellRange = ExcelHelper.CreateRange(cells, range);

            cellRange.UnMerge();
            workbook.Save(outputPath);

            return $"Range {range} unmerged. Output: {outputPath}";
        });
    }

    /// <summary>
    ///     Gets all merged cell ranges from the worksheet.
    /// </summary>
    /// <param name="path">Excel file path.</param>
    /// <param name="sheetIndex">Worksheet index (0-based).</param>
    /// <returns>JSON string containing all merged ranges.</returns>
    private Task<string> GetMergedCellsAsync(string path, int sheetIndex)
    {
        return Task.Run(() =>
        {
            using var workbook = new Workbook(path);
            var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);
            var mergedCells = worksheet.Cells.MergedCells;

            if (mergedCells == null || mergedCells.Count == 0)
            {
                var emptyResult = new
                {
                    count = 0,
                    worksheetName = worksheet.Name,
                    items = Array.Empty<object>(),
                    message = "No merged cells found"
                };
                return JsonSerializer.Serialize(emptyResult, new JsonSerializerOptions { WriteIndented = true });
            }

            var mergedList = new List<object>();
            for (var i = 0; i < mergedCells.Count; i++)
            {
                var mergedCellObj = mergedCells[i];

                if (mergedCellObj is CellArea cellArea)
                {
                    var startCellName = CellsHelper.CellIndexToName(cellArea.StartRow, cellArea.StartColumn);
                    var endCellName = CellsHelper.CellIndexToName(cellArea.EndRow, cellArea.EndColumn);
                    var rangeName = $"{startCellName}:{endCellName}";

                    var cell = worksheet.Cells[cellArea.StartRow, cellArea.StartColumn];
                    var cellValue = cell.Value?.ToString() ?? "(empty)";

                    mergedList.Add(new
                    {
                        index = i,
                        range = rangeName,
                        startCell = startCellName,
                        endCell = endCellName,
                        rowCount = cellArea.EndRow - cellArea.StartRow + 1,
                        columnCount = cellArea.EndColumn - cellArea.StartColumn + 1,
                        value = cellValue
                    });
                }
            }

            var result = new
            {
                count = mergedList.Count,
                worksheetName = worksheet.Name,
                items = mergedList
            };

            return JsonSerializer.Serialize(result, new JsonSerializerOptions { WriteIndented = true });
        });
    }
}