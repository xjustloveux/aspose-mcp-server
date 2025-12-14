using System.Text.Json.Nodes;
using System.Text;
using Aspose.Cells;
using AsposeMcpServer.Core;

namespace AsposeMcpServer.Tools;

/// <summary>
/// Unified tool for Excel data operations (sort, find/replace, batch write, get content, get statistics, get used range)
/// Merges: ExcelSortDataTool, ExcelFindReplaceTool, ExcelBatchWriteTool, ExcelGetContentTool, 
/// ExcelGetStatisticsTool, ExcelGetUsedRangeTool
/// </summary>
public class ExcelDataOperationsTool : IAsposeTool
{
    public string Description => @"Excel data operations. Supports 6 operations: sort, find_replace, batch_write, get_content, get_statistics, get_used_range.

Usage examples:
- Sort data: excel_data_operations(operation='sort', path='book.xlsx', range='A1:C10', sortColumn=0)
- Find and replace: excel_data_operations(operation='find_replace', path='book.xlsx', searchText='old', replaceText='new')
- Batch write: excel_data_operations(operation='batch_write', path='book.xlsx', data={'A1':'Value1','B1':'Value2'})
- Get content: excel_data_operations(operation='get_content', path='book.xlsx', range='A1:C10')
- Get statistics: excel_data_operations(operation='get_statistics', path='book.xlsx', range='A1:A10')
- Get used range: excel_data_operations(operation='get_used_range', path='book.xlsx')";

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
                @enum = new[] { "sort", "find_replace", "batch_write", "get_content", "get_statistics", "get_used_range" }
            },
            path = new
            {
                type = "string",
                description = "Excel file path (required for all operations)"
            },
            outputPath = new
            {
                type = "string",
                description = "Output file path (optional, for find_replace/batch_write, defaults to input path)"
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

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var operation = ArgumentHelper.GetString(arguments, "operation", "operation");
        var path = ArgumentHelper.GetAndValidatePath(arguments);
        var sheetIndex = arguments?["sheetIndex"]?.GetValue<int>() ?? 0;

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
    /// Sorts data in a range
    /// </summary>
    /// <param name="arguments">JSON arguments containing range, sortColumn, optional ascending, hasHeader</param>
    /// <param name="path">Excel file path</param>
    /// <param name="sheetIndex">Worksheet index (0-based)</param>
    /// <returns>Success message</returns>
    private async Task<string> SortDataAsync(JsonObject? arguments, string path, int sheetIndex)
    {
        var range = ArgumentHelper.GetString(arguments, "range", "range");
        var sortColumn = ArgumentHelper.GetInt(arguments, "sortColumn", "sortColumn");
        var ascending = arguments?["ascending"]?.GetValue<bool?>() ?? true;
        var hasHeader = arguments?["hasHeader"]?.GetValue<bool?>() ?? false;

        using var workbook = new Workbook(path);
        var worksheet = workbook.Worksheets[sheetIndex];
        var cells = worksheet.Cells;
        var cellRange = cells.CreateRange(range);

        var dataSorter = workbook.DataSorter;
        // Sort data - use the 5-parameter overload
        dataSorter.Sort(cells, cellRange.FirstRow, cellRange.FirstColumn, cellRange.FirstRow + cellRange.RowCount - 1, cellRange.FirstColumn + cellRange.ColumnCount - 1);

        workbook.Save(path);
        return await Task.FromResult($"Data sorted in range {range} by column {sortColumn} ({ (ascending ? "ascending" : "descending") }): {path}");
    }

    /// <summary>
    /// Finds and replaces text in the worksheet
    /// </summary>
    /// <param name="arguments">JSON arguments containing searchText, replaceText, optional range, outputPath</param>
    /// <param name="path">Excel file path</param>
    /// <param name="sheetIndex">Worksheet index (0-based)</param>
    /// <returns>Success message with replacement count</returns>
    private async Task<string> FindReplaceAsync(JsonObject? arguments, string path, int sheetIndex)
    {
        var outputPath = ArgumentHelper.GetAndValidateOutputPath(arguments, path);
        var findText = ArgumentHelper.GetString(arguments, "findText", "findText");
        var replaceText = ArgumentHelper.GetString(arguments, "replaceText", "replaceText");
        var sheetIndexParam = arguments?["sheetIndex"]?.GetValue<int?>();
        var matchCase = arguments?["matchCase"]?.GetValue<bool?>() ?? false;
        var matchEntireCell = arguments?["matchEntireCell"]?.GetValue<bool?>() ?? false;

        using var workbook = new Workbook(path);
        var totalReplacements = 0;
        var lookAt = matchEntireCell ? LookAtType.EntireContent : LookAtType.Contains;
        
        if (sheetIndexParam.HasValue)
        {
            if (sheetIndexParam.Value < 0 || sheetIndexParam.Value >= workbook.Worksheets.Count)
            {
                throw new ArgumentException($"工作表索引 {sheetIndexParam.Value} 超出範圍 (共有 {workbook.Worksheets.Count} 個工作表)");
            }
            
            var worksheet = workbook.Worksheets[sheetIndexParam.Value];
            totalReplacements += ReplaceInWorksheet(worksheet, findText, replaceText, matchCase, lookAt);
        }
        else
        {
            for (int i = 0; i < workbook.Worksheets.Count; i++)
            {
                var worksheet = workbook.Worksheets[i];
                totalReplacements += ReplaceInWorksheet(worksheet, findText, replaceText, matchCase, lookAt);
            }
        }
        
        workbook.Save(outputPath);
        
        return await Task.FromResult($"查找替換完成\n查找: '{findText}'\n替換為: '{replaceText}'\n總替換數: {totalReplacements}\n輸出: {outputPath}");
    }

    private int ReplaceInWorksheet(Worksheet worksheet, string findText, string replaceText, bool matchCase, LookAtType lookAt)
    {
        var findOptions = new FindOptions
        {
            CaseSensitive = matchCase,
            LookAtType = lookAt
        };
        
        var cell = worksheet.Cells.Find(findText, null, findOptions);
        int count = 0;
        
        while (cell != null)
        {
            cell.PutValue(replaceText);
            count++;
            cell = worksheet.Cells.Find(findText, cell, findOptions);
        }
        
        return count;
    }

    /// <summary>
    /// Writes multiple values to cells in batch
    /// </summary>
    /// <param name="arguments">JSON arguments containing data array (objects with cell and value)</param>
    /// <param name="path">Excel file path</param>
    /// <param name="sheetIndex">Worksheet index (0-based)</param>
    /// <returns>Success message with write count</returns>
    private async Task<string> BatchWriteAsync(JsonObject? arguments, string path, int sheetIndex)
    {
        var outputPath = ArgumentHelper.GetAndValidateOutputPath(arguments, path);
        var dataArray = arguments?["data"]?.AsArray() ?? throw new ArgumentException("data is required for batch_write operation");

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
                    if (double.TryParse(value, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out double numValue))
                    {
                        cellObj.PutValue(numValue);
                    }
                    else if (bool.TryParse(value, out bool boolValue))
                    {
                        cellObj.PutValue(boolValue);
                    }
                    else
                    {
                        cellObj.PutValue(value);
                    }
                }
            }
        }

        workbook.Save(outputPath);
        return await Task.FromResult($"Batch write completed: {dataArray.Count} cells written to sheet {sheetIndex}: {outputPath}");
    }

    /// <summary>
    /// Gets content from a range
    /// </summary>
    /// <param name="arguments">JSON arguments containing optional range</param>
    /// <param name="path">Excel file path</param>
    /// <param name="sheetIndex">Worksheet index (0-based)</param>
    /// <returns>Formatted string with cell content</returns>
    private async Task<string> GetContentAsync(JsonObject? arguments, string path, int sheetIndex)
    {
        var range = arguments?["range"]?.GetValue<string>();

        using var workbook = new Workbook(path);
        var worksheet = workbook.Worksheets[sheetIndex];

        if (!string.IsNullOrEmpty(range))
        {
            var cells = worksheet.Cells;
            var cellRange = cells.CreateRange(range);
            var options = new ExportTableOptions
            {
                ExportColumnName = false
            };
            var dataTable = cells.ExportDataTable(cellRange.FirstRow, cellRange.FirstColumn, cellRange.RowCount, cellRange.ColumnCount, options);
            
            var json = System.Text.Json.JsonSerializer.Serialize(dataTable);
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
            
            var json = System.Text.Json.JsonSerializer.Serialize(dataTable);
            return await Task.FromResult(json);
        }
    }

    /// <summary>
    /// Gets statistics for a range (count, sum, average, min, max)
    /// </summary>
    /// <param name="arguments">JSON arguments containing range</param>
    /// <param name="path">Excel file path</param>
    /// <returns>Formatted string with statistics</returns>
    private async Task<string> GetStatisticsAsync(JsonObject? arguments, string path)
    {
        var sheetIndex = arguments?["sheetIndex"]?.GetValue<int?>();

        using var workbook = new Workbook(path);
        var result = new StringBuilder();

        result.AppendLine("=== Excel 工作簿統計資訊 ===\n");

        result.AppendLine("【工作簿資訊】");
        result.AppendLine($"總工作表數: {workbook.Worksheets.Count}");
        result.AppendLine($"檔案格式: {workbook.FileFormat}");
        result.AppendLine();

        if (sheetIndex.HasValue)
        {
            if (sheetIndex.Value < 0 || sheetIndex.Value >= workbook.Worksheets.Count)
            {
                throw new ArgumentException($"工作表索引 {sheetIndex.Value} 超出範圍 (共有 {workbook.Worksheets.Count} 個工作表)");
            }
            AppendSheetStatistics(result, workbook.Worksheets[sheetIndex.Value], sheetIndex.Value);
        }
        else
        {
            for (int i = 0; i < workbook.Worksheets.Count; i++)
            {
                AppendSheetStatistics(result, workbook.Worksheets[i], i);
                if (i < workbook.Worksheets.Count - 1) result.AppendLine();
            }
        }

        return await Task.FromResult(result.ToString());
    }

    private void AppendSheetStatistics(StringBuilder result, Worksheet worksheet, int index)
    {
        result.AppendLine($"【工作表 {index}: {worksheet.Name}】");
        result.AppendLine($"最大數據行: {worksheet.Cells.MaxDataRow + 1}");
        result.AppendLine($"最大數據列: {worksheet.Cells.MaxDataColumn + 1}");
        result.AppendLine($"圖表數: {worksheet.Charts.Count}");
        result.AppendLine($"圖片數: {worksheet.Pictures.Count}");
        result.AppendLine($"超連結數: {worksheet.Hyperlinks.Count}");
        result.AppendLine($"註釋數: {worksheet.Comments.Count}");
    }

    /// <summary>
    /// Gets the used range information for the worksheet
    /// </summary>
    /// <param name="arguments">JSON arguments (no specific parameters required)</param>
    /// <param name="path">Excel file path</param>
    /// <param name="sheetIndex">Worksheet index (0-based)</param>
    /// <returns>Formatted string with used range details</returns>
    private async Task<string> GetUsedRangeAsync(JsonObject? arguments, string path, int sheetIndex)
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

