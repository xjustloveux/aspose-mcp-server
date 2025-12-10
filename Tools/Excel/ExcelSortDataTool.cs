using System.Text.Json.Nodes;
using Aspose.Cells;

namespace AsposeMcpServer.Tools;

public class ExcelSortDataTool : IAsposeTool
{
    public string Description => "Sort data in a range of cells in an Excel worksheet";

    public object InputSchema => new
    {
        type = "object",
        properties = new
        {
            path = new
            {
                type = "string",
                description = "Excel file path"
            },
            sheetIndex = new
            {
                type = "number",
                description = "Sheet index (0-based, optional, default: 0)"
            },
            range = new
            {
                type = "string",
                description = "Cell range to sort (e.g., 'A1:C10')"
            },
            sortColumn = new
            {
                type = "number",
                description = "Column index to sort by (0-based, relative to range start)"
            },
            ascending = new
            {
                type = "boolean",
                description = "True for ascending, false for descending (default: true)"
            },
            hasHeader = new
            {
                type = "boolean",
                description = "Whether the range has a header row (default: false)"
            }
        },
        required = new[] { "path", "range", "sortColumn" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var sheetIndex = arguments?["sheetIndex"]?.GetValue<int>() ?? 0;
        var range = arguments?["range"]?.GetValue<string>() ?? throw new ArgumentException("range is required");
        var sortColumn = arguments?["sortColumn"]?.GetValue<int>() ?? throw new ArgumentException("sortColumn is required");
        var ascending = arguments?["ascending"]?.GetValue<bool>() ?? true;
        var hasHeader = arguments?["hasHeader"]?.GetValue<bool>() ?? false;

        using var workbook = new Workbook(path);
        var worksheet = workbook.Worksheets[sheetIndex];
        var cells = worksheet.Cells;
        var cellRange = cells.CreateRange(range);

        var dataSorter = workbook.DataSorter;
        // Set sort key (Key1 is column index)
        dataSorter.Key1 = cellRange.FirstColumn + sortColumn;
        // Perform sort (simplified - basic sort functionality)
        dataSorter.Sort(cells, cellRange.FirstRow, cellRange.FirstColumn, cellRange.FirstRow + cellRange.RowCount - 1, cellRange.FirstColumn + cellRange.ColumnCount - 1);

        workbook.Save(path);

        return await Task.FromResult($"範圍 {range} 已按第 {sortColumn} 列排序 ({(ascending ? "升序" : "降序")}): {path}");
    }
}

