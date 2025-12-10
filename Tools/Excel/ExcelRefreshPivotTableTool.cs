using System.Text.Json.Nodes;
using Aspose.Cells;

namespace AsposeMcpServer.Tools;

public class ExcelRefreshPivotTableTool : IAsposeTool
{
    public string Description => "Refresh pivot table data in an Excel worksheet";

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
            outputPath = new
            {
                type = "string",
                description = "Output file path (if not provided, overwrites input)"
            },
            sheetIndex = new
            {
                type = "number",
                description = "Sheet index (0-based, optional, default: 0)"
            },
            pivotTableIndex = new
            {
                type = "number",
                description = "Pivot table index to refresh (0-based, optional, if not provided refreshes all pivot tables)"
            }
        },
        required = new[] { "path" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var outputPath = arguments?["outputPath"]?.GetValue<string>() ?? path;
        var sheetIndex = arguments?["sheetIndex"]?.GetValue<int>() ?? 0;
        var pivotTableIndex = arguments?["pivotTableIndex"]?.GetValue<int?>();

        using var workbook = new Workbook(path);
        
        if (sheetIndex < 0 || sheetIndex >= workbook.Worksheets.Count)
        {
            throw new ArgumentException($"工作表索引 {sheetIndex} 超出範圍 (共有 {workbook.Worksheets.Count} 個工作表)");
        }

        var worksheet = workbook.Worksheets[sheetIndex];
        var pivotTables = worksheet.PivotTables;
        
        if (pivotTables.Count == 0)
        {
            throw new InvalidOperationException($"工作表 '{worksheet.Name}' 中未找到樞紐表");
        }

        int refreshedCount = 0;

        if (pivotTableIndex.HasValue)
        {
            if (pivotTableIndex.Value < 0 || pivotTableIndex.Value >= pivotTables.Count)
            {
                throw new ArgumentException($"樞紐表索引 {pivotTableIndex.Value} 超出範圍 (工作表共有 {pivotTables.Count} 個樞紐表)");
            }
            
            pivotTables[pivotTableIndex.Value].CalculateData();
            refreshedCount = 1;
        }
        else
        {
            // Refresh all pivot tables
            foreach (var pivotTable in pivotTables)
            {
                pivotTable.CalculateData();
                refreshedCount++;
            }
        }

        workbook.Save(outputPath);

        return await Task.FromResult($"成功刷新 {refreshedCount} 個樞紐表\n工作表: {worksheet.Name}\n輸出: {outputPath}");
    }
}

