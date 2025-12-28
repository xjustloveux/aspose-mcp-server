using System.Data;
using System.Globalization;
using System.Text.Json;
using System.Text.Json.Nodes;
using Aspose.Cells;
using AsposeMcpServer.Core;

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
                description =
                    "Data for batch_write. Supports two formats: (1) Array: [{\"cell\":\"A1\",\"value\":\"val1\"},...] (2) Object: {\"A1\":\"val1\",\"B1\":\"val2\"} - more compact",
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
        var outputPath = ArgumentHelper.GetAndValidateOutputPath(arguments, path);
        var sheetIndex = ArgumentHelper.GetInt(arguments, "sheetIndex", 0);

        return operation.ToLower() switch
        {
            "sort" => await SortDataAsync(path, outputPath, sheetIndex, arguments),
            "find_replace" => await FindReplaceAsync(path, outputPath, arguments),
            "batch_write" => await BatchWriteAsync(path, outputPath, sheetIndex, arguments),
            "get_content" => await GetContentAsync(path, sheetIndex, arguments),
            "get_statistics" => await GetStatisticsAsync(path, arguments),
            "get_used_range" => await GetUsedRangeAsync(path, sheetIndex),
            _ => throw new ArgumentException($"Unknown operation: {operation}")
        };
    }

    /// <summary>
    ///     Sorts data in a range
    /// </summary>
    /// <param name="path">Excel file path</param>
    /// <param name="outputPath">Output file path</param>
    /// <param name="sheetIndex">Worksheet index (0-based)</param>
    /// <param name="arguments">JSON arguments containing range, sortColumn, optional ascending, hasHeader</param>
    /// <returns>Success message</returns>
    private Task<string> SortDataAsync(string path, string outputPath, int sheetIndex, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var range = ArgumentHelper.GetString(arguments, "range");
            var sortColumn = ArgumentHelper.GetInt(arguments, "sortColumn");
            var ascending = ArgumentHelper.GetBool(arguments, "ascending", true);
            var hasHeader = ArgumentHelper.GetBool(arguments, "hasHeader", false);

            try
            {
                using var workbook = new Workbook(path);
                var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);
                var cells = worksheet.Cells;

                var cellRange = ExcelHelper.CreateRange(cells, range);

                var rows = new List<List<object?>>();
                var startRow = hasHeader ? cellRange.FirstRow + 1 : cellRange.FirstRow;

                if (hasHeader)
                {
                    var headerRow = new List<object?>();
                    for (var col = cellRange.FirstColumn; col < cellRange.FirstColumn + cellRange.ColumnCount; col++)
                        headerRow.Add(cells[cellRange.FirstRow, col].Value);
                    rows.Add(headerRow);
                }

                for (var row = startRow; row < cellRange.FirstRow + cellRange.RowCount; row++)
                {
                    var rowData = new List<object?>();
                    for (var col = cellRange.FirstColumn; col < cellRange.FirstColumn + cellRange.ColumnCount; col++)
                        rowData.Add(cells[row, col].Value);
                    rows.Add(rowData);
                }

                var dataRows = hasHeader ? rows.Skip(1).ToList() : rows;
                dataRows.Sort((a, b) =>
                {
                    var aVal = a[sortColumn];
                    var bVal = b[sortColumn];

                    if (aVal == null && bVal == null) return 0;
                    if (aVal == null) return ascending ? -1 : 1;
                    if (bVal == null) return ascending ? 1 : -1;

                    var comparison = Comparer<object>.Default.Compare(aVal, bVal);
                    return ascending ? comparison : -comparison;
                });

                if (hasHeader)
                {
                    rows = [rows[0]];
                    rows.AddRange(dataRows);
                }
                else
                {
                    rows = dataRows;
                }

                for (var i = 0; i < rows.Count; i++)
                {
                    var rowData = rows[i];
                    var targetRow = cellRange.FirstRow + i;
                    for (var j = 0; j < rowData.Count; j++)
                        cells[targetRow, cellRange.FirstColumn + j].Value = rowData[j];
                }

                workbook.Save(outputPath);
                return
                    $"Sorted range {range} by column {sortColumn} ({(ascending ? "ascending" : "descending")}). Output: {outputPath}";
            }
            catch (CellsException ex)
            {
                throw new ArgumentException($"Excel operation failed for range '{range}': {ex.Message}");
            }
        });
    }

    /// <summary>
    ///     Finds and replaces text in the worksheet
    /// </summary>
    /// <param name="path">Excel file path</param>
    /// <param name="outputPath">Output file path</param>
    /// <param name="arguments">JSON arguments containing searchText, replaceText, optional range, outputPath</param>
    /// <returns>Success message with replacement count</returns>
    private Task<string> FindReplaceAsync(string path, string outputPath, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var findText = ArgumentHelper.GetString(arguments, "findText");
            var replaceText = ArgumentHelper.GetString(arguments, "replaceText");
            var sheetIndexParam = ArgumentHelper.GetIntNullable(arguments, "sheetIndex");
            var matchCase = ArgumentHelper.GetBool(arguments, "matchCase", false);
            var matchEntireCell = ArgumentHelper.GetBool(arguments, "matchEntireCell", false);

            try
            {
                using var workbook = new Workbook(path);
                var totalReplacements = 0;
                var lookAt = matchEntireCell ? LookAtType.EntireContent : LookAtType.Contains;

                if (sheetIndexParam.HasValue)
                {
                    var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndexParam.Value);
                    totalReplacements += ReplaceInWorksheet(worksheet, findText, replaceText, matchCase, lookAt);
                }
                else
                {
                    foreach (var worksheet in workbook.Worksheets)
                        totalReplacements += ReplaceInWorksheet(worksheet, findText, replaceText, matchCase, lookAt);
                }

                workbook.Save(outputPath);
                return
                    $"Replaced '{findText}' with '{replaceText}' ({totalReplacements} replacements). Output: {outputPath}";
            }
            catch (CellsException ex)
            {
                throw new ArgumentException($"Excel operation failed: {ex.Message}");
            }
        });
    }

    /// <summary>
    ///     Replaces text in a worksheet
    /// </summary>
    /// <param name="worksheet">Worksheet to replace text in</param>
    /// <param name="findText">Text to find</param>
    /// <param name="replaceText">Replacement text</param>
    /// <param name="matchCase">Whether to match case</param>
    /// <param name="lookAt">Look at type (entire content or contains)</param>
    /// <returns>Number of replacements made</returns>
    private int ReplaceInWorksheet(Worksheet worksheet, string findText, string replaceText, bool matchCase,
        LookAtType lookAt)
    {
        var findOptions = new FindOptions
        {
            CaseSensitive = matchCase,
            LookAtType = lookAt
        };

        var replacedCells = new HashSet<string>();
        var cell = worksheet.Cells.Find(findText, null, findOptions);
        var count = 0;

        while (cell != null)
        {
            var cellName = cell.Name;
            if (replacedCells.Contains(cellName))
                break;

            if (lookAt == LookAtType.EntireContent)
            {
                cell.PutValue(replaceText);
            }
            else
            {
                var currentValue = cell.StringValue ?? "";
                var comparison = matchCase ? StringComparison.Ordinal : StringComparison.OrdinalIgnoreCase;
                var newValue = currentValue.Replace(findText, replaceText, comparison);
                cell.PutValue(newValue);
            }

            replacedCells.Add(cellName);
            count++;
            cell = worksheet.Cells.Find(findText, cell, findOptions);
        }

        return count;
    }

    /// <summary>
    ///     Writes multiple values to cells in batch
    /// </summary>
    /// <param name="path">Excel file path</param>
    /// <param name="outputPath">Output file path</param>
    /// <param name="sheetIndex">Worksheet index (0-based)</param>
    /// <param name="arguments">JSON arguments containing data array (objects with cell and value)</param>
    /// <returns>Success message with write count</returns>
    private Task<string> BatchWriteAsync(string path, string outputPath, int sheetIndex, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            try
            {
                using var workbook = new Workbook(path);
                var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);
                var writeCount = 0;

                if (arguments != null && arguments.TryGetPropertyValue("data", out var dataNode))
                {
                    if (dataNode is JsonArray dataArray)
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
                                    ExcelHelper.SetCellValue(cellObj, value);
                                    writeCount++;
                                }
                            }
                        }
                    else if (dataNode is JsonObject dataObject)
                        foreach (var kvp in dataObject)
                        {
                            var cell = kvp.Key;
                            var value = kvp.Value?.GetValue<string>() ?? "";
                            if (!string.IsNullOrEmpty(cell))
                            {
                                var cellObj = worksheet.Cells[cell];
                                ExcelHelper.SetCellValue(cellObj, value);
                                writeCount++;
                            }
                        }
                }

                workbook.Save(outputPath);
                return
                    $"Batch write completed ({writeCount} cells written to sheet {sheetIndex}). Output: {outputPath}";
            }
            catch (CellsException ex)
            {
                throw new ArgumentException($"Excel operation failed: {ex.Message}");
            }
        });
    }

    /// <summary>
    ///     Gets content from a range
    /// </summary>
    /// <param name="path">Excel file path</param>
    /// <param name="sheetIndex">Worksheet index (0-based)</param>
    /// <param name="arguments">JSON arguments containing optional range</param>
    /// <returns>Formatted string with cell content</returns>
    private Task<string> GetContentAsync(string path, int sheetIndex, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var range = ArgumentHelper.GetStringNullable(arguments, "range");

            try
            {
                using var workbook = new Workbook(path);
                var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);
                var cells = worksheet.Cells;

                var jsonOptions = new JsonSerializerOptions { WriteIndented = false };

                if (!string.IsNullOrEmpty(range))
                {
                    var cellRange = ExcelHelper.CreateRange(cells, range);
                    var options = new ExportTableOptions { ExportColumnName = false };
                    var dataTable = cells.ExportDataTable(cellRange.FirstRow, cellRange.FirstColumn,
                        cellRange.RowCount, cellRange.ColumnCount, options);

                    var rows = ConvertDataTableToList(dataTable);
                    return JsonSerializer.Serialize(rows, jsonOptions);
                }
                else
                {
                    var maxRow = cells.MaxDataRow;
                    var maxCol = cells.MaxDataColumn;
                    var options = new ExportTableOptions { ExportColumnName = false };
                    var dataTable = cells.ExportDataTable(0, 0, maxRow + 1, maxCol + 1, options);

                    var rows = ConvertDataTableToList(dataTable);
                    return JsonSerializer.Serialize(rows, jsonOptions);
                }
            }
            catch (CellsException ex)
            {
                throw new ArgumentException($"Excel operation failed: {ex.Message}");
            }
        });
    }

    /// <summary>
    ///     Converts a DataTable to a list of dictionaries for JSON serialization
    /// </summary>
    /// <param name="dataTable">DataTable to convert</param>
    /// <returns>List of dictionaries representing rows</returns>
    private static List<Dictionary<string, object?>> ConvertDataTableToList(DataTable dataTable)
    {
        var rows = new List<Dictionary<string, object?>>();
        foreach (DataRow row in dataTable.Rows)
        {
            var rowDict = new Dictionary<string, object?>();
            foreach (DataColumn column in dataTable.Columns)
            {
                var value = row[column];
                rowDict[column.ColumnName] = value == DBNull.Value ? null : value;
            }

            rows.Add(rowDict);
        }

        return rows;
    }

    /// <summary>
    ///     Gets statistics for a range (count, sum, average, min, max)
    /// </summary>
    /// <param name="path">Excel file path</param>
    /// <param name="arguments">JSON arguments containing optional range and sheetIndex</param>
    /// <returns>JSON formatted string with statistics</returns>
    private Task<string> GetStatisticsAsync(string path, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var sheetIndex = ArgumentHelper.GetIntNullable(arguments, "sheetIndex");
            var range = ArgumentHelper.GetStringNullable(arguments, "range");

            try
            {
                using var workbook = new Workbook(path);
                var worksheets = new List<object>();

                if (sheetIndex.HasValue)
                {
                    var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex.Value);
                    worksheets.Add(GetSheetStatistics(worksheet, sheetIndex.Value, range));
                }
                else
                {
                    for (var i = 0; i < workbook.Worksheets.Count; i++)
                        worksheets.Add(GetSheetStatistics(workbook.Worksheets[i], i, range));
                }

                var result = new
                {
                    totalWorksheets = workbook.Worksheets.Count,
                    fileFormat = workbook.FileFormat.ToString(),
                    worksheets
                };

                return JsonSerializer.Serialize(result, new JsonSerializerOptions { WriteIndented = true });
            }
            catch (CellsException ex)
            {
                throw new ArgumentException($"Excel operation failed: {ex.Message}");
            }
        });
    }

    /// <summary>
    ///     Gets statistics for a worksheet
    /// </summary>
    /// <param name="worksheet">Worksheet to get statistics from</param>
    /// <param name="index">Worksheet index</param>
    /// <param name="range">Optional range to calculate detailed statistics for</param>
    /// <returns>Dictionary containing worksheet statistics</returns>
    private object GetSheetStatistics(Worksheet worksheet, int index, string? range)
    {
        var baseStats = new Dictionary<string, object>
        {
            ["index"] = index,
            ["name"] = worksheet.Name,
            ["maxDataRow"] = worksheet.Cells.MaxDataRow + 1,
            ["maxDataColumn"] = worksheet.Cells.MaxDataColumn + 1,
            ["chartsCount"] = worksheet.Charts.Count,
            ["picturesCount"] = worksheet.Pictures.Count,
            ["hyperlinksCount"] = worksheet.Hyperlinks.Count,
            ["commentsCount"] = worksheet.Comments.Count
        };

        // If range is specified, calculate detailed statistics for that range
        if (!string.IsNullOrEmpty(range))
            try
            {
                var cellRange = ExcelHelper.CreateRange(worksheet.Cells, range);
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

                var rangeStats = new Dictionary<string, object>
                {
                    ["range"] = range,
                    ["totalCells"] = cellRange.RowCount * cellRange.ColumnCount,
                    ["numericCells"] = numericValues.Count,
                    ["nonNumericCells"] = nonNumericCount,
                    ["emptyCells"] = emptyCount
                };

                if (numericValues.Count > 0)
                {
                    numericValues.Sort();
                    rangeStats["sum"] = Math.Round(numericValues.Sum(), 2);
                    rangeStats["average"] = Math.Round(numericValues.Sum() / numericValues.Count, 2);
                    rangeStats["min"] = Math.Round(numericValues[0], 2);
                    rangeStats["max"] = Math.Round(numericValues[^1], 2);
                    rangeStats["count"] = numericValues.Count;
                }

                baseStats["rangeStatistics"] = rangeStats;
            }
            catch (Exception ex)
            {
                baseStats["rangeStatisticsError"] = ex.Message;
            }

        return baseStats;
    }

    /// <summary>
    ///     Gets the used range information for the worksheet
    /// </summary>
    /// <param name="path">Excel file path</param>
    /// <param name="sheetIndex">Worksheet index (0-based)</param>
    /// <returns>JSON formatted string with used range details</returns>
    private Task<string> GetUsedRangeAsync(string path, int sheetIndex)
    {
        return Task.Run(() =>
        {
            try
            {
                using var workbook = new Workbook(path);
                var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);
                var cells = worksheet.Cells;

                string? rangeAddress = null;
                if (cells.MaxDataRow >= cells.MinDataRow && cells.MaxDataColumn >= cells.MinDataColumn)
                {
                    var firstCell = CellsHelper.CellIndexToName(cells.MinDataRow, cells.MinDataColumn);
                    var lastCell = CellsHelper.CellIndexToName(cells.MaxDataRow, cells.MaxDataColumn);
                    rangeAddress = $"{firstCell}:{lastCell}";
                }

                var result = new
                {
                    worksheetName = worksheet.Name,
                    sheetIndex,
                    firstRow = cells.MinDataRow,
                    lastRow = cells.MaxDataRow,
                    firstColumn = cells.MinDataColumn,
                    lastColumn = cells.MaxDataColumn,
                    range = rangeAddress
                };

                return JsonSerializer.Serialize(result, new JsonSerializerOptions { WriteIndented = true });
            }
            catch (CellsException ex)
            {
                throw new ArgumentException($"Excel operation failed: {ex.Message}");
            }
        });
    }
}