using System.Globalization;
using System.Text;
using System.Text.Json;
using System.Text.Json.Nodes;
using Aspose.Cells;
using AsposeMcpServer.Core;
using Range = Aspose.Cells.Range;

namespace AsposeMcpServer.Tools.Excel;

/// <summary>
///     Unified tool for Excel data operations (sort, find/replace, batch write, get content, get statistics, get used
///     range)
///     Merges: ExcelSortDataTool, ExcelFindReplaceTool, ExcelBatchWriteTool, ExcelGetContentTool,
///     ExcelGetStatisticsTool, ExcelGetUsedRangeTool
/// </summary>
public class ExcelDataOperationsTool : IAsposeTool
{
    /// <summary>
    ///     Gets the description of the tool and its usage examples
    /// </summary>
    public string Description =>
        @"Excel data operations. Supports 6 operations: sort, find_replace, batch_write, get_content, get_statistics, get_used_range.

Usage examples:
- Sort data: excel_data_operations(operation='sort', path='book.xlsx', range='A1:C10', sortColumn=0)
- Find and replace: excel_data_operations(operation='find_replace', path='book.xlsx', searchText='old', replaceText='new')
- Batch write: excel_data_operations(operation='batch_write', path='book.xlsx', data={'A1':'Value1','B1':'Value2'})
- Get content: excel_data_operations(operation='get_content', path='book.xlsx', range='A1:C10')
- Get statistics: excel_data_operations(operation='get_statistics', path='book.xlsx', range='A1:A10')
- Get used range: excel_data_operations(operation='get_used_range', path='book.xlsx')";

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
- 'sort': Sort data in range (required params: path, range, sortColumn)
- 'find_replace': Find and replace text (required params: path, searchText, replaceText)
- 'batch_write': Write multiple values at once (required params: path, data)
- 'get_content': Get cell content (required params: path)
- 'get_statistics': Get statistics for range (required params: path, range)
- 'get_used_range': Get used range info (required params: path)",
                @enum = new[]
                    { "sort", "find_replace", "batch_write", "get_content", "get_statistics", "get_used_range" }
            },
            path = new
            {
                type = "string",
                description = "Excel file path (required for all operations)"
            },
            outputPath = new
            {
                type = "string",
                description =
                    "Output file path (optional, for sort/find_replace/batch_write operations, defaults to input path)"
            },
            sheetIndex = new
            {
                type = "number",
                description = "Sheet index (0-based, optional, default: 0)"
            },
            range = new
            {
                type = "string",
                description = "Cell range (e.g., 'A1:C10', required for sort, optional for get_content)"
            },
            sortColumn = new
            {
                type = "number",
                description = "Column index to sort by (0-based, relative to range start, required for sort)"
            },
            ascending = new
            {
                type = "boolean",
                description = "True for ascending, false for descending (optional, for sort, default: true)"
            },
            hasHeader = new
            {
                type = "boolean",
                description = "Whether the range has a header row (optional, for sort, default: false)"
            },
            findText = new
            {
                type = "string",
                description = "Text to find (required for find_replace)"
            },
            replaceText = new
            {
                type = "string",
                description = "Text to replace with (required for find_replace)"
            },
            matchCase = new
            {
                type = "boolean",
                description = "Match case (optional, for find_replace, default: false)"
            },
            matchEntireCell = new
            {
                type = "boolean",
                description = "Match entire cell content (optional, for find_replace, default: false)"
            },
            data = new
            {
                type = "array",
                description = "Array of objects with 'cell' and 'value' properties (required for batch_write)",
                items = new
                {
                    type = "object",
                    properties = new
                    {
                        cell = new { type = "string" },
                        value = new { type = "string" }
                    },
                    required = new[] { "cell", "value" }
                }
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
            "sort" => await SortDataAsync(arguments, path, sheetIndex),
            "find_replace" => await FindReplaceAsync(arguments, path, sheetIndex),
            "batch_write" => await BatchWriteAsync(arguments, path, sheetIndex),
            "get_content" => await GetContentAsync(arguments, path, sheetIndex),
            "get_statistics" => await GetStatisticsAsync(arguments, path),
            "get_used_range" => await GetUsedRangeAsync(arguments, path, sheetIndex),
            _ => throw new ArgumentException($"Unknown operation: {operation}")
        };
    }

    /// <summary>
    ///     Sorts data in a range
    /// </summary>
    /// <param name="arguments">JSON arguments containing range, sortColumn, optional ascending, hasHeader</param>
    /// <param name="path">Excel file path</param>
    /// <param name="sheetIndex">Worksheet index (0-based)</param>
    /// <returns>Success message</returns>
    private async Task<string> SortDataAsync(JsonObject? arguments, string path, int sheetIndex)
    {
        var outputPath = ArgumentHelper.GetAndValidateOutputPath(arguments, path);
        var range = ArgumentHelper.GetString(arguments, "range");
        var sortColumn = ArgumentHelper.GetInt(arguments, "sortColumn");
        var ascending = ArgumentHelper.GetBool(arguments, "ascending");
        _ = ArgumentHelper.GetBool(arguments, "hasHeader", false);

        using var workbook = new Workbook(path);
        var worksheet = workbook.Worksheets[sheetIndex];
        var cells = worksheet.Cells;

        var cellRange = ExcelHelper.CreateRange(cells, range);

        var dataSorter = workbook.DataSorter;
        // Sort data - use the 5-parameter overload
        dataSorter.Sort(cells, cellRange.FirstRow, cellRange.FirstColumn, cellRange.FirstRow + cellRange.RowCount - 1,
            cellRange.FirstColumn + cellRange.ColumnCount - 1);

        workbook.Save(outputPath);
        return await Task.FromResult(
            $"Data sorted in range {range} by column {sortColumn} ({(ascending ? "ascending" : "descending")}): {outputPath}");
    }

    /// <summary>
    ///     Finds and replaces text in the worksheet
    /// </summary>
    /// <param name="arguments">JSON arguments containing searchText, replaceText, optional range, outputPath</param>
    /// <param name="path">Excel file path</param>
    /// <param name="_">Unused parameter</param>
    /// <returns>Success message with replacement count</returns>
    private async Task<string> FindReplaceAsync(JsonObject? arguments, string path, int _)
    {
        var outputPath = ArgumentHelper.GetAndValidateOutputPath(arguments, path);
        var findText = ArgumentHelper.GetString(arguments, "findText");
        var replaceText = ArgumentHelper.GetString(arguments, "replaceText");
        var sheetIndexParam = ArgumentHelper.GetIntNullable(arguments, "sheetIndex");
        var matchCase = ArgumentHelper.GetBool(arguments, "matchCase", false);
        var matchEntireCell = ArgumentHelper.GetBool(arguments, "matchEntireCell", false);

        using var workbook = new Workbook(path);
        var totalReplacements = 0;
        var lookAt = matchEntireCell ? LookAtType.EntireContent : LookAtType.Contains;

        if (sheetIndexParam.HasValue)
        {
            if (sheetIndexParam.Value < 0 || sheetIndexParam.Value >= workbook.Worksheets.Count)
                throw new ArgumentException(
                    $"Worksheet index {sheetIndexParam.Value} is out of range (workbook has {workbook.Worksheets.Count} worksheets)");

            var worksheet = workbook.Worksheets[sheetIndexParam.Value];
            totalReplacements += ReplaceInWorksheet(worksheet, findText, replaceText, matchCase, lookAt);
        }
        else
        {
            foreach (var worksheet in workbook.Worksheets)
                totalReplacements += ReplaceInWorksheet(worksheet, findText, replaceText, matchCase, lookAt);
        }

        workbook.Save(outputPath);

        return await Task.FromResult(
            $"Find and replace completed\nFind: '{findText}'\nReplace with: '{replaceText}'\nTotal replacements: {totalReplacements}\nOutput: {outputPath}");
    }

    private int ReplaceInWorksheet(Worksheet worksheet, string findText, string replaceText, bool matchCase,
        LookAtType lookAt)
    {
        var findOptions = new FindOptions
        {
            CaseSensitive = matchCase,
            LookAtType = lookAt
        };

        var cell = worksheet.Cells.Find(findText, null, findOptions);
        var count = 0;

        while (cell != null)
        {
            cell.PutValue(replaceText);
            count++;
            cell = worksheet.Cells.Find(findText, cell, findOptions);
        }

        return count;
    }

    /// <summary>
    ///     Writes multiple values to cells in batch
    /// </summary>
    /// <param name="arguments">JSON arguments containing data array (objects with cell and value)</param>
    /// <param name="path">Excel file path</param>
    /// <param name="sheetIndex">Worksheet index (0-based)</param>
    /// <returns>Success message with write count</returns>
    private async Task<string> BatchWriteAsync(JsonObject? arguments, string path, int sheetIndex)
    {
        var outputPath = ArgumentHelper.GetAndValidateOutputPath(arguments, path);
        var dataArray = ArgumentHelper.GetArray(arguments, "data");

        using var workbook = new Workbook(path);
        var worksheet = workbook.Worksheets[sheetIndex];

        foreach (var item in dataArray)
        {
            var itemObj = item?.AsObject();
            if (itemObj != null)
            {
                var cell = itemObj["cell"]?.GetValue<string>();
                var value = itemObj["value"]?.GetValue<string>() ?? "";
                if (!string.IsNullOrEmpty(cell))
                {
                    var cellObj = worksheet.Cells[cell];

                    // Parse value as number, boolean, or string
                    // Ensures conditional formatting and charts can correctly identify numeric values
                    if (double.TryParse(value, NumberStyles.Any, CultureInfo.InvariantCulture, out var numValue))
                        cellObj.PutValue(numValue);
                    else if (bool.TryParse(value, out var boolValue))
                        cellObj.PutValue(boolValue);
                    else
                        cellObj.PutValue(value);
                }
            }
        }

        workbook.Save(outputPath);
        return await Task.FromResult(
            $"Batch write completed: {dataArray.Count} cells written to sheet {sheetIndex}: {outputPath}");
    }

    /// <summary>
    ///     Gets content from a range
    /// </summary>
    /// <param name="arguments">JSON arguments containing optional range</param>
    /// <param name="path">Excel file path</param>
    /// <param name="sheetIndex">Worksheet index (0-based)</param>
    /// <returns>Formatted string with cell content</returns>
    private async Task<string> GetContentAsync(JsonObject? arguments, string path, int sheetIndex)
    {
        var range = ArgumentHelper.GetStringNullable(arguments, "range");

        using var workbook = new Workbook(path);
        var worksheet = workbook.Worksheets[sheetIndex];

        if (!string.IsNullOrEmpty(range))
        {
            var cells = worksheet.Cells;

            Range cellRange;
            try
            {
                cellRange = ExcelHelper.CreateRange(cells, range);
            }
            catch (Exception ex)
            {
                throw new ArgumentException(
                    $"Invalid range format: '{range}'. Range exceeds Excel limits (valid range: A1:XFD1048576). Error: {ex.Message}");
            }

            var options = new ExportTableOptions
            {
                ExportColumnName = false
            };
            var dataTable = cells.ExportDataTable(cellRange.FirstRow, cellRange.FirstColumn, cellRange.RowCount,
                cellRange.ColumnCount, options);

            var json = JsonSerializer.Serialize(dataTable);
            return await Task.FromResult(json);
        }
        else
        {
            var maxRow = worksheet.Cells.MaxDataRow;
            var maxCol = worksheet.Cells.MaxDataColumn;
            var options = new ExportTableOptions
            {
                ExportColumnName = false
            };
            var dataTable = worksheet.Cells.ExportDataTable(0, 0, maxRow + 1, maxCol + 1, options);

            var json = JsonSerializer.Serialize(dataTable);
            return await Task.FromResult(json);
        }
    }

    /// <summary>
    ///     Gets statistics for a range (count, sum, average, min, max)
    /// </summary>
    /// <param name="arguments">JSON arguments containing optional range and sheetIndex</param>
    /// <param name="path">Excel file path</param>
    /// <returns>Formatted string with statistics</returns>
    private async Task<string> GetStatisticsAsync(JsonObject? arguments, string path)
    {
        var sheetIndex = ArgumentHelper.GetIntNullable(arguments, "sheetIndex");
        var range = ArgumentHelper.GetStringNullable(arguments, "range");

        using var workbook = new Workbook(path);
        var result = new StringBuilder();

        result.AppendLine("=== Excel Workbook Statistics ===\n");

        result.AppendLine("[Workbook Information]");
        result.AppendLine($"Total worksheets: {workbook.Worksheets.Count}");
        result.AppendLine($"File format: {workbook.FileFormat}");
        result.AppendLine();

        if (sheetIndex.HasValue)
        {
            if (sheetIndex.Value < 0 || sheetIndex.Value >= workbook.Worksheets.Count)
                throw new ArgumentException(
                    $"Worksheet index {sheetIndex.Value} is out of range (workbook has {workbook.Worksheets.Count} worksheets)");
            AppendSheetStatistics(result, workbook.Worksheets[sheetIndex.Value], sheetIndex.Value, range);
        }
        else
        {
            for (var i = 0; i < workbook.Worksheets.Count; i++)
            {
                AppendSheetStatistics(result, workbook.Worksheets[i], i, range);
                if (i < workbook.Worksheets.Count - 1) result.AppendLine();
            }
        }

        return await Task.FromResult(result.ToString());
    }

    private void AppendSheetStatistics(StringBuilder result, Worksheet worksheet, int index, string? range)
    {
        result.AppendLine($"[Worksheet {index}: {worksheet.Name}]");
        result.AppendLine($"Max data row: {worksheet.Cells.MaxDataRow + 1}");
        result.AppendLine($"Max data column: {worksheet.Cells.MaxDataColumn + 1}");
        result.AppendLine($"Charts count: {worksheet.Charts.Count}");
        result.AppendLine($"Pictures count: {worksheet.Pictures.Count}");
        result.AppendLine($"Hyperlinks count: {worksheet.Hyperlinks.Count}");
        result.AppendLine($"Comments count: {worksheet.Comments.Count}");

        // If range is specified, calculate detailed statistics for that range
        if (!string.IsNullOrEmpty(range))
            try
            {
                var cellRange = ExcelHelper.CreateRange(worksheet.Cells, range);
                result.AppendLine();
                result.AppendLine($"[Range Statistics for '{range}']");

                var numericValues = new List<double>();
                var nonNumericCount = 0;
                var emptyCount = 0;

                for (var row = cellRange.FirstRow; row < cellRange.FirstRow + cellRange.RowCount; row++)
                for (var col = cellRange.FirstColumn; col < cellRange.FirstColumn + cellRange.ColumnCount; col++)
                {
                    var cell = worksheet.Cells[row, col];
                    var value = cell.Value;

                    if (value == null || (value is string str && string.IsNullOrWhiteSpace(str)))
                        emptyCount++;
                    else if (value is double || value is int || value is float || value is decimal)
                        numericValues.Add(Convert.ToDouble(value));
                    else if (double.TryParse(value.ToString(), NumberStyles.Any,
                                 CultureInfo.InvariantCulture, out var numValue))
                        numericValues.Add(numValue);
                    else
                        nonNumericCount++;
                }

                result.AppendLine($"Total cells: {cellRange.RowCount * cellRange.ColumnCount}");
                result.AppendLine($"Numeric cells: {numericValues.Count}");
                result.AppendLine($"Non-numeric cells: {nonNumericCount}");
                result.AppendLine($"Empty cells: {emptyCount}");

                if (numericValues.Count > 0)
                {
                    numericValues.Sort();
                    var sum = numericValues.Sum();
                    var average = sum / numericValues.Count;
                    var min = numericValues[0];
                    var max = numericValues[numericValues.Count - 1];

                    result.AppendLine($"Sum: {sum:F2}");
                    result.AppendLine($"Average: {average:F2}");
                    result.AppendLine($"Min: {min:F2}");
                    result.AppendLine($"Max: {max:F2}");
                    result.AppendLine($"Count: {numericValues.Count}");
                }
                else
                {
                    result.AppendLine("No numeric values found in the specified range.");
                }
            }
            catch (Exception ex)
            {
                result.AppendLine($"Could not calculate range statistics: {ex.Message}");
            }
    }

    /// <summary>
    ///     Gets the used range information for the worksheet
    /// </summary>
    /// <param name="_">Unused parameter</param>
    /// <param name="path">Excel file path</param>
    /// <param name="sheetIndex">Worksheet index (0-based)</param>
    /// <returns>Formatted string with used range details</returns>
    private async Task<string> GetUsedRangeAsync(JsonObject? _, string path, int sheetIndex)
    {
        using var workbook = new Workbook(path);
        var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);
        var cells = worksheet.Cells;

        var sb = new StringBuilder();
        sb.AppendLine($"Used Range for Sheet '{worksheet.Name}':");
        sb.AppendLine($"  First Row: {cells.MinDataRow}");
        sb.AppendLine($"  Last Row: {cells.MaxDataRow}");
        sb.AppendLine($"  First Column: {cells.MinDataColumn}");
        sb.AppendLine($"  Last Column: {cells.MaxDataColumn}");

        if (cells.MaxDataRow >= cells.MinDataRow && cells.MaxDataColumn >= cells.MinDataColumn)
        {
            var firstCell = CellsHelper.CellIndexToName(cells.MinDataRow, cells.MinDataColumn);
            var lastCell = CellsHelper.CellIndexToName(cells.MaxDataRow, cells.MaxDataColumn);
            sb.AppendLine($"  Range: {firstCell}:{lastCell}");
        }

        return await Task.FromResult(sb.ToString());
    }
}