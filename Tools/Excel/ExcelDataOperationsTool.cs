using System.ComponentModel;
using System.Data;
using System.Globalization;
using System.Text.Json;
using System.Text.Json.Nodes;
using Aspose.Cells;
using AsposeMcpServer.Core.Helpers;
using AsposeMcpServer.Core.Session;
using ModelContextProtocol.Server;

namespace AsposeMcpServer.Tools.Excel;

/// <summary>
///     Unified tool for Excel data operations (sort, find/replace, batch write, get content, get statistics, get used
///     range)
/// </summary>
[McpServerToolType]
public class ExcelDataOperationsTool
{
    /// <summary>
    ///     Document session manager for in-memory editing support.
    /// </summary>
    private readonly DocumentSessionManager? _sessionManager;

    /// <summary>
    ///     Initializes a new instance of the <see cref="ExcelDataOperationsTool" /> class.
    /// </summary>
    /// <param name="sessionManager">Optional session manager for in-memory document editing.</param>
    public ExcelDataOperationsTool(DocumentSessionManager? sessionManager = null)
    {
        _sessionManager = sessionManager;
    }

    [McpServerTool(Name = "excel_data_operations")]
    [Description(
        @"Excel data operations. Supports 6 operations: sort, find_replace, batch_write, get_content, get_statistics, get_used_range.

Usage examples:
- Sort data: excel_data_operations(operation='sort', path='book.xlsx', range='A1:C10', sortColumn=0)
- Find and replace: excel_data_operations(operation='find_replace', path='book.xlsx', findText='old', replaceText='new')
- Batch write: excel_data_operations(operation='batch_write', path='book.xlsx', data=[{cell:'A1',value:'Value1'},{cell:'B1',value:'Value2'}])
- Get content: excel_data_operations(operation='get_content', path='book.xlsx', range='A1:C10')
- Get statistics: excel_data_operations(operation='get_statistics', path='book.xlsx', range='A1:A10')
- Get used range: excel_data_operations(operation='get_used_range', path='book.xlsx')")]
    public string Execute(
        [Description("Operation: sort, find_replace, batch_write, get_content, get_statistics, get_used_range")]
        string operation,
        [Description("Excel file path (required if no sessionId)")]
        string? path = null,
        [Description("Session ID for in-memory editing")]
        string? sessionId = null,
        [Description("Output file path (file mode only)")]
        string? outputPath = null,
        [Description("Sheet index (0-based, default: 0)")]
        int sheetIndex = 0,
        [Description("Cell range (e.g., 'A1:C10', required for sort, optional for get_content)")]
        string? range = null,
        [Description("Column index to sort by (0-based, relative to range start, required for sort)")]
        int sortColumn = 0,
        [Description("True for ascending, false for descending (default: true)")]
        bool ascending = true,
        [Description("Whether the range has a header row (default: false)")]
        bool hasHeader = false,
        [Description("Text to find (required for find_replace)")]
        string? findText = null,
        [Description("Text to replace with (required for find_replace)")]
        string? replaceText = null,
        [Description("Match case (default: false)")]
        bool matchCase = false,
        [Description("Match entire cell content (default: false)")]
        bool matchEntireCell = false,
        [Description("Data for batch_write: [{cell:'A1',value:'val1'},...] or JSON object {A1:'val1',B1:'val2'}")]
        JsonNode? data = null)
    {
        using var ctx = DocumentContext<Workbook>.Create(_sessionManager, sessionId, path);

        return operation.ToLower() switch
        {
            "sort" => SortData(ctx, outputPath, sheetIndex, range, sortColumn, ascending, hasHeader),
            "find_replace" => FindReplace(ctx, outputPath, sheetIndex, findText, replaceText, matchCase,
                matchEntireCell),
            "batch_write" => BatchWrite(ctx, outputPath, sheetIndex, data),
            "get_content" => GetContent(ctx, sheetIndex, range),
            "get_statistics" => GetStatistics(ctx, sheetIndex, range),
            "get_used_range" => GetUsedRange(ctx, sheetIndex),
            _ => throw new ArgumentException($"Unknown operation: {operation}")
        };
    }

    /// <summary>
    ///     Sorts data in a range.
    /// </summary>
    /// <param name="ctx">The document context containing the workbook.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="sheetIndex">The worksheet index.</param>
    /// <param name="range">The cell range to sort.</param>
    /// <param name="sortColumn">The column index to sort by (0-based, relative to range start).</param>
    /// <param name="ascending">True for ascending order, false for descending.</param>
    /// <param name="hasHeader">Whether the range has a header row.</param>
    /// <returns>A message indicating the result of the sort operation.</returns>
    /// <exception cref="ArgumentException">Thrown when range is not provided or Excel operation fails.</exception>
    private static string SortData(DocumentContext<Workbook> ctx, string? outputPath, int sheetIndex,
        string? range, int sortColumn, bool ascending, bool hasHeader)
    {
        if (string.IsNullOrEmpty(range))
            throw new ArgumentException("range is required for sort operation");

        try
        {
            var workbook = ctx.Document;
            var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);
            var cells = worksheet.Cells;

            var cellRange = ExcelHelper.CreateRange(cells, range);

            List<List<object?>> rows = [];
            var startRow = hasHeader ? cellRange.FirstRow + 1 : cellRange.FirstRow;

            if (hasHeader)
            {
                List<object?> headerRow = [];
                for (var col = cellRange.FirstColumn; col < cellRange.FirstColumn + cellRange.ColumnCount; col++)
                    headerRow.Add(cells[cellRange.FirstRow, col].Value);
                rows.Add(headerRow);
            }

            for (var row = startRow; row < cellRange.FirstRow + cellRange.RowCount; row++)
            {
                List<object?> rowData = [];
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

            ctx.Save(outputPath);
            return
                $"Sorted range {range} by column {sortColumn} ({(ascending ? "ascending" : "descending")}). {ctx.GetOutputMessage(outputPath)}";
        }
        catch (CellsException ex)
        {
            throw new ArgumentException($"Excel operation failed for range '{range}': {ex.Message}");
        }
    }

    /// <summary>
    ///     Finds and replaces text in the worksheet.
    /// </summary>
    /// <param name="ctx">The document context containing the workbook.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="sheetIndex">The worksheet index, or null to search all worksheets.</param>
    /// <param name="findText">The text to find.</param>
    /// <param name="replaceText">The text to replace with.</param>
    /// <param name="matchCase">Whether to match case.</param>
    /// <param name="matchEntireCell">Whether to match entire cell content.</param>
    /// <returns>A message indicating the number of replacements made.</returns>
    /// <exception cref="ArgumentException">Thrown when findText or replaceText is not provided, or Excel operation fails.</exception>
    private static string FindReplace(DocumentContext<Workbook> ctx, string? outputPath, int? sheetIndex,
        string? findText, string? replaceText, bool matchCase, bool matchEntireCell)
    {
        if (string.IsNullOrEmpty(findText))
            throw new ArgumentException("findText is required for find_replace operation");
        if (replaceText == null)
            throw new ArgumentException("replaceText is required for find_replace operation");

        try
        {
            var workbook = ctx.Document;
            var totalReplacements = 0;
            var lookAt = matchEntireCell ? LookAtType.EntireContent : LookAtType.Contains;

            if (sheetIndex.HasValue)
            {
                var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex.Value);
                totalReplacements += ReplaceInWorksheet(worksheet, findText, replaceText, matchCase, lookAt);
            }
            else
            {
                foreach (var worksheet in workbook.Worksheets)
                    totalReplacements += ReplaceInWorksheet(worksheet, findText, replaceText, matchCase, lookAt);
            }

            ctx.Save(outputPath);
            return
                $"Replaced '{findText}' with '{replaceText}' ({totalReplacements} replacements). {ctx.GetOutputMessage(outputPath)}";
        }
        catch (CellsException ex)
        {
            throw new ArgumentException($"Excel operation failed: {ex.Message}");
        }
    }

    /// <summary>
    ///     Replaces text in a worksheet.
    /// </summary>
    /// <param name="worksheet">The worksheet to search in.</param>
    /// <param name="findText">The text to find.</param>
    /// <param name="replaceText">The text to replace with.</param>
    /// <param name="matchCase">Whether to match case.</param>
    /// <param name="lookAt">The type of match (entire content or contains).</param>
    /// <returns>The number of replacements made.</returns>
    private static int ReplaceInWorksheet(Worksheet worksheet, string findText, string replaceText, bool matchCase,
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
    ///     Writes multiple values to cells in batch.
    /// </summary>
    /// <param name="ctx">The document context containing the workbook.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="sheetIndex">The worksheet index.</param>
    /// <param name="data">The data to write as JSON array or object.</param>
    /// <returns>A message indicating the number of cells written.</returns>
    /// <exception cref="ArgumentException">Thrown when Excel operation fails.</exception>
    private static string BatchWrite(DocumentContext<Workbook> ctx, string? outputPath, int sheetIndex, JsonNode? data)
    {
        try
        {
            var workbook = ctx.Document;
            var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);
            var writeCount = 0;

            if (data != null)
            {
                if (data is JsonArray dataArray)
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
                else if (data is JsonObject dataObject)
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

            ctx.Save(outputPath);
            return
                $"Batch write completed ({writeCount} cells written to sheet {sheetIndex}). {ctx.GetOutputMessage(outputPath)}";
        }
        catch (CellsException ex)
        {
            throw new ArgumentException($"Excel operation failed: {ex.Message}");
        }
    }

    /// <summary>
    ///     Gets content from a range.
    /// </summary>
    /// <param name="ctx">The document context containing the workbook.</param>
    /// <param name="sheetIndex">The worksheet index.</param>
    /// <param name="range">The cell range to get content from, or null for all data.</param>
    /// <returns>A JSON string containing the range content.</returns>
    /// <exception cref="ArgumentException">Thrown when Excel operation fails.</exception>
    private static string GetContent(DocumentContext<Workbook> ctx, int sheetIndex, string? range)
    {
        try
        {
            var workbook = ctx.Document;
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
    }

    /// <summary>
    ///     Converts a DataTable to a list of dictionaries for JSON serialization.
    /// </summary>
    /// <param name="dataTable">The DataTable to convert.</param>
    /// <returns>A list of dictionaries representing the table rows.</returns>
    private static List<Dictionary<string, object?>> ConvertDataTableToList(DataTable dataTable)
    {
        List<Dictionary<string, object?>> rows = [];
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
    ///     Gets statistics for a range (count, sum, average, min, max).
    /// </summary>
    /// <param name="ctx">The document context containing the workbook.</param>
    /// <param name="sheetIndex">The worksheet index, or null for all worksheets.</param>
    /// <param name="range">The cell range to calculate statistics for.</param>
    /// <returns>A JSON string containing the statistics.</returns>
    /// <exception cref="ArgumentException">Thrown when Excel operation fails.</exception>
    private static string GetStatistics(DocumentContext<Workbook> ctx, int? sheetIndex, string? range)
    {
        try
        {
            var workbook = ctx.Document;
            List<object> worksheets = [];

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
    }

    /// <summary>
    ///     Gets statistics for a worksheet.
    /// </summary>
    /// <param name="worksheet">The worksheet to get statistics for.</param>
    /// <param name="index">The worksheet index.</param>
    /// <param name="range">The cell range to calculate statistics for, or null for basic sheet info.</param>
    /// <returns>An object containing the worksheet statistics.</returns>
    private static object GetSheetStatistics(Worksheet worksheet, int index, string? range)
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

        if (!string.IsNullOrEmpty(range))
            try
            {
                var cellRange = ExcelHelper.CreateRange(worksheet.Cells, range);
                List<double> numericValues = [];
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
    ///     Gets the used range information for the worksheet.
    /// </summary>
    /// <param name="ctx">The document context containing the workbook.</param>
    /// <param name="sheetIndex">The worksheet index.</param>
    /// <returns>A JSON string containing the used range information.</returns>
    /// <exception cref="ArgumentException">Thrown when Excel operation fails.</exception>
    private static string GetUsedRange(DocumentContext<Workbook> ctx, int sheetIndex)
    {
        try
        {
            var workbook = ctx.Document;
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
    }
}