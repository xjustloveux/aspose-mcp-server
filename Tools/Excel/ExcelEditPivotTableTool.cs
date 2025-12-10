using System.Text.Json.Nodes;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeMcpServer.Tools;

public class ExcelEditPivotTableTool : IAsposeTool
{
    public string Description => "Edit an existing pivot table in an Excel worksheet";

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
                description = "Pivot table index to edit (0-based)"
            },
            name = new
            {
                type = "string",
                description = "New name for pivot table (optional)"
            },
            refreshData = new
            {
                type = "boolean",
                description = "Refresh pivot table data (default: false)"
            }
        },
        required = new[] { "path", "pivotTableIndex" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var outputPath = arguments?["outputPath"]?.GetValue<string>() ?? path;
        var sheetIndex = arguments?["sheetIndex"]?.GetValue<int>() ?? 0;
        var pivotTableIndex = arguments?["pivotTableIndex"]?.GetValue<int>() ?? throw new ArgumentException("pivotTableIndex is required");
        var name = arguments?["name"]?.GetValue<string>();
        var refreshData = arguments?["refreshData"]?.GetValue<bool>() ?? false;

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
        var changes = new List<string>();

        // Update name
        if (!string.IsNullOrEmpty(name))
        {
            pivotTable.Name = name;
            changes.Add($"名稱: {name}");
        }

        // Refresh data
        if (refreshData)
        {
            pivotTable.CalculateData();
            changes.Add("數據已刷新");
        }

        workbook.Save(outputPath);

        var result = $"成功編輯樞紐表 #{pivotTableIndex}\n";
        if (changes.Count > 0)
        {
            result += "變更:\n";
            foreach (var change in changes)
            {
                result += $"  - {change}\n";
            }
        }
        else
        {
            result += "無變更。\n";
        }
        result += $"輸出: {outputPath}";

        return await Task.FromResult(result);
    }
}

