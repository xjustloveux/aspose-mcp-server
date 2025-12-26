using System.Text.Json;
using System.Text.Json.Nodes;
using Aspose.Cells;
using AsposeMcpServer.Core;

namespace AsposeMcpServer.Tools.Excel;

/// <summary>
///     Unified tool for managing Excel merged cells (merge/unmerge/get)
///     Merges: ExcelMergeCellsTool, ExcelGetMergedCellsTool
/// </summary>
public class ExcelMergeCellsTool : IAsposeTool
{
    /// <summary>
    ///     Gets the description of the tool and its usage examples
    /// </summary>
    public string Description => @"Manage Excel merged cells. Supports 3 operations: merge, unmerge, get.

Usage examples:
- Merge cells: excel_merge_cells(operation='merge', path='book.xlsx', range='A1:C1')
- Unmerge cells: excel_merge_cells(operation='unmerge', path='book.xlsx', range='A1:C1')
- Get merged cells: excel_merge_cells(operation='get', path='book.xlsx')";

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
- 'merge': Merge cells (required params: path, range)
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
                description = "Cell range to merge/unmerge (e.g., 'A1:C3', required for merge/unmerge)"
            },
            outputPath = new
            {
                type = "string",
                description = "Output file path (optional, for merge/unmerge operations, defaults to input path)"
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
            "merge" => await MergeCellsAsync(path, outputPath, sheetIndex, arguments),
            "unmerge" => await UnmergeCellsAsync(path, outputPath, sheetIndex, arguments),
            "get" => await GetMergedCellsAsync(path, sheetIndex),
            _ => throw new ArgumentException($"Unknown operation: {operation}")
        };
    }

    /// <summary>
    ///     Merges cells in a range
    /// </summary>
    /// <param name="path">Excel file path</param>
    /// <param name="outputPath">Output file path</param>
    /// <param name="sheetIndex">Worksheet index (0-based)</param>
    /// <param name="arguments">JSON arguments containing range</param>
    /// <returns>Success message with merged range</returns>
    private Task<string> MergeCellsAsync(string path, string outputPath, int sheetIndex, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var range = ArgumentHelper.GetString(arguments, "range");

            using var workbook = new Workbook(path);
            var worksheet = workbook.Worksheets[sheetIndex];
            var cells = worksheet.Cells;

            var cellRange = ExcelHelper.CreateRange(cells, range);

            cellRange.Merge();
            workbook.Save(outputPath);
            return $"Range {range} merged. Output: {outputPath}";
        });
    }

    /// <summary>
    ///     Unmerges cells in a range
    /// </summary>
    /// <param name="path">Excel file path</param>
    /// <param name="outputPath">Output file path</param>
    /// <param name="sheetIndex">Worksheet index (0-based)</param>
    /// <param name="arguments">JSON arguments containing range</param>
    /// <returns>Success message</returns>
    private Task<string> UnmergeCellsAsync(string path, string outputPath, int sheetIndex, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var range = ArgumentHelper.GetString(arguments, "range");

            using var workbook = new Workbook(path);
            var worksheet = workbook.Worksheets[sheetIndex];
            var cells = worksheet.Cells;

            var cellRange = ExcelHelper.CreateRange(cells, range);

            cellRange.UnMerge();
            workbook.Save(outputPath);
            return $"Range {range} unmerged. Output: {outputPath}";
        });
    }

    /// <summary>
    ///     Gets all merged cell ranges from the worksheet
    /// </summary>
    /// <param name="path">Excel file path</param>
    /// <param name="sheetIndex">Worksheet index (0-based)</param>
    /// <returns>JSON string with all merged ranges</returns>
    private Task<string> GetMergedCellsAsync(string path, int sheetIndex)
    {
        return Task.Run(() =>
        {
            using var workbook = new Workbook(path);
            var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);
            var mergedCells = worksheet.Cells.MergedCells;
            if (mergedCells == null)
                throw new InvalidOperationException($"Unable to get merged cells information: {worksheet.Name}");

            if (mergedCells.Count == 0)
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
                try
                {
                    string? mergedCellName = null;
                    int startRow = -1, startCol = -1, endRow = -1, endCol = -1;
                    string? cellValue = null;
                    string? error = null;

                    try
                    {
                        var mergedCellObj = mergedCells[i];

                        if (mergedCellObj is CellArea cellArea)
                        {
                            startRow = cellArea.StartRow;
                            startCol = cellArea.StartColumn;
                            endRow = cellArea.EndRow;
                            endCol = cellArea.EndColumn;

                            var startCellName = CellsHelper.CellIndexToName(startRow, startCol);
                            var endCellName = CellsHelper.CellIndexToName(endRow, endCol);
                            mergedCellName = $"{startCellName}:{endCellName}";
                        }
                        else
                        {
                            mergedCellName = mergedCells[i]?.ToString();

                            if (!string.IsNullOrWhiteSpace(mergedCellName))
                            {
                                if (mergedCellName.Contains("CellArea(") && mergedCellName.Contains("["))
                                {
                                    var bracketStart = mergedCellName.IndexOf('[');
                                    if (bracketStart > 0)
                                    {
                                        var bracketPart = mergedCellName.Substring(bracketStart + 1);
                                        bracketPart = bracketPart.TrimEnd(']');
                                        var indices = bracketPart.Split(',');
                                        if (indices.Length >= 4)
                                            if (int.TryParse(indices[0].Trim(), out startRow) &&
                                                int.TryParse(indices[1].Trim(), out startCol) &&
                                                int.TryParse(indices[2].Trim(), out endRow) &&
                                                int.TryParse(indices[3].Trim(), out endCol))
                                            {
                                                var startCellName = CellsHelper.CellIndexToName(startRow, startCol);
                                                var endCellName = CellsHelper.CellIndexToName(endRow, endCol);
                                                mergedCellName = $"{startCellName}:{endCellName}";
                                            }
                                    }

                                    if (startRow < 0)
                                    {
                                        var startIdx = mergedCellName.IndexOf('(') + 1;
                                        var endIdx = mergedCellName.IndexOf(')', startIdx);
                                        if (startIdx > 0 && endIdx > startIdx)
                                        {
                                            var rangePart = mergedCellName.Substring(startIdx, endIdx - startIdx);
                                            mergedCellName = rangePart;
                                        }
                                    }
                                }

                                if (startRow < 0 && mergedCellName.Contains(':'))
                                {
                                    var parts = mergedCellName.Split(':');
                                    if (parts.Length == 2)
                                    {
                                        var startCell = parts[0].Trim();
                                        var endCell = parts[1].Trim();

                                        try
                                        {
                                            CellsHelper.CellNameToIndex(startCell, out startRow, out startCol);
                                            CellsHelper.CellNameToIndex(endCell, out endRow, out endCol);
                                        }
                                        catch (Exception ex)
                                        {
                                            Console.Error.WriteLine(
                                                $"[WARN] Failed to parse cell names '{startCell}' or '{endCell}': {ex.Message}");
                                        }
                                    }
                                }
                            }
                        }

                        if (startRow >= 0 && startCol >= 0)
                            try
                            {
                                var cell = worksheet.Cells[startRow, startCol];
                                cellValue = cell.Value?.ToString() ?? "(empty)";
                            }
                            catch (Exception ex)
                            {
                                cellValue = "(unable to read)";
                                Console.Error.WriteLine(
                                    $"[WARN] Failed to read cell value at [{startRow}, {startCol}]: {ex.Message}");
                            }
                    }
                    catch (Exception ex)
                    {
                        error = $"Unable to read merged cell information - {ex.Message}";
                    }

                    mergedList.Add(new
                    {
                        index = i,
                        range = mergedCellName ?? "unknown",
                        rowCount = endRow >= 0 && startRow >= 0 ? endRow - startRow + 1 : (int?)null,
                        columnCount = endCol >= 0 && startCol >= 0 ? endCol - startCol + 1 : (int?)null,
                        value = cellValue,
                        error
                    });
                }
                catch (Exception ex)
                {
                    mergedList.Add(new
                    {
                        index = i,
                        range = (string?)null,
                        rowCount = (int?)null,
                        columnCount = (int?)null,
                        value = (string?)null,
                        error = $"Unable to read merged cell information - {ex.Message}"
                    });
                }

            var result = new
            {
                count = mergedCells.Count,
                worksheetName = worksheet.Name,
                items = mergedList
            };

            return JsonSerializer.Serialize(result, new JsonSerializerOptions { WriteIndented = true });
        });
    }
}