using System.Text.Json.Nodes;
using Aspose.Cells;

namespace AsposeMcpServer.Tools;

public class ExcelDeletePivotTableTool : IAsposeTool
{
    public string Description => "Delete a pivot table from an Excel worksheet";

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
            pivotTableIndex = new
            {
                type = "number",
                description = "Pivot table index to delete (0-based)"
            }
        },
        required = new[] { "path", "pivotTableIndex" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var sheetIndex = arguments?["sheetIndex"]?.GetValue<int>() ?? 0;
        var pivotTableIndex = arguments?["pivotTableIndex"]?.GetValue<int>() ?? throw new ArgumentException("pivotTableIndex is required");

        using var workbook = new Workbook(path);
        
        if (sheetIndex < 0 || sheetIndex >= workbook.Worksheets.Count)
        {
            throw new ArgumentException($"工作表索引 {sheetIndex} 超出範圍 (共有 {workbook.Worksheets.Count} 個工作表)");
        }

        var worksheet = workbook.Worksheets[sheetIndex];
        var pivotTables = worksheet.PivotTables;
        
        if (pivotTableIndex < 0 || pivotTableIndex >= pivotTables.Count)
        {
            throw new ArgumentException($"樞紐表索引 {pivotTableIndex} 超出範圍 (工作表共有 {pivotTables.Count} 個樞紐表)");
        }

        var pivotTable = pivotTables[pivotTableIndex];
        var pivotTableName = pivotTable.Name ?? $"樞紐表 {pivotTableIndex}";
        
        pivotTables.RemoveAt(pivotTableIndex);
        workbook.Save(path);
        
        var remainingCount = pivotTables.Count;
        
        return await Task.FromResult($"成功刪除樞紐表 #{pivotTableIndex} ({pivotTableName})\n工作表剩餘樞紐表數: {remainingCount}\n輸出: {path}");
    }
}

