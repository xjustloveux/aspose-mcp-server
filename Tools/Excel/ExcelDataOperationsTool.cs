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
    public string Description => "Excel data operations: sort, find/replace, batch write, get content, get statistics, get used range";

    public object InputSchema => new
    {
        type = "object",
        properties = new
        {
            operation = new
            {
                type = "string",
                description = "Operation to perform: 'sort', 'find_replace', 'batch_write', 'get_content', 'get_statistics', 'get_used_range'",
                @enum = new[] { "sort", "find_replace", "batch_write", "get_content", "get_statistics", "get_used_range" }
            },
            path = new
            {
                type = "string",
                description = "Excel file path"
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
        var operation = arguments?["operation"]?.GetValue<string>() ?? throw new ArgumentException("operation is required");
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        SecurityHelper.ValidateFilePath(path, "path");
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

    private async Task<string> SortDataAsync(JsonObject? arguments, string path, int sheetIndex)
    {
        var range = arguments?["range"]?.GetValue<string>() ?? throw new ArgumentException("range is required for sort operation");
        var sortColumn = arguments?["sortColumn"]?.GetValue<int>() ?? throw new ArgumentException("sortColumn is required for sort operation");
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

    private async Task<string> FindReplaceAsync(JsonObject? arguments, string path, int sheetIndex)
    {
        var outputPath = arguments?["outputPath"]?.GetValue<string>() ?? path;
        SecurityHelper.ValidateFilePath(outputPath, "outputPath");
        var findText = arguments?["findText"]?.GetValue<string>() ?? throw new ArgumentException("findText is required for find_replace operation");
        var replaceText = arguments?["replaceText"]?.GetValue<string>() ?? throw new ArgumentException("replaceText is required for find_replace operation");
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

    private async Task<string> BatchWriteAsync(JsonObject? arguments, string path, int sheetIndex)
    {
        var outputPath = arguments?["outputPath"]?.GetValue<string>() ?? path;
        SecurityHelper.ValidateFilePath(outputPath, "outputPath");
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
                    
                    // 嘗試將值解析為數字，如果是數字則設定為數字類型，否則設定為字符串
                    // 這樣可以確保條件格式和圖表能正確識別數字值
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

