using System.Text.Json.Nodes;
using System.Text;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeMcpServer.Tools;

public class ExcelGetPivotTablesTool : IAsposeTool
{
    public string Description => "Get all pivot tables information from an Excel worksheet";

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
            }
        },
        required = new[] { "path" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var sheetIndex = arguments?["sheetIndex"]?.GetValue<int>() ?? 0;

        using var workbook = new Workbook(path);
        
        if (sheetIndex < 0 || sheetIndex >= workbook.Worksheets.Count)
        {
            throw new ArgumentException($"工作表索引 {sheetIndex} 超出範圍 (共有 {workbook.Worksheets.Count} 個工作表)");
        }

        var worksheet = workbook.Worksheets[sheetIndex];
        var pivotTables = worksheet.PivotTables;
        var result = new StringBuilder();

        result.AppendLine($"=== 工作表 '{worksheet.Name}' 的樞紐表資訊 ===\n");
        result.AppendLine($"總樞紐表數: {pivotTables.Count}\n");

        if (pivotTables.Count == 0)
        {
            result.AppendLine("未找到樞紐表");
            return await Task.FromResult(result.ToString());
        }

        for (int i = 0; i < pivotTables.Count; i++)
        {
            var pivotTable = pivotTables[i];
            result.AppendLine($"【樞紐表 {i}】");
            result.AppendLine($"名稱: {pivotTable.Name ?? "(無名稱)"}");
            result.AppendLine($"數據源: {pivotTable.DataSource}");
            var dataBodyRange = pivotTable.DataBodyRange;
            if (dataBodyRange.StartRow >= 0)
            {
                result.AppendLine($"位置: 行 {dataBodyRange.StartRow}-{dataBodyRange.EndRow}, 列 {dataBodyRange.StartColumn}-{dataBodyRange.EndColumn}");
            }
            else
            {
                result.AppendLine($"位置: 未知");
            }
            
            if (pivotTable.RowFields != null && pivotTable.RowFields.Count > 0)
            {
                result.AppendLine($"行欄位數: {pivotTable.RowFields.Count}");
            }
            
            if (pivotTable.ColumnFields != null && pivotTable.ColumnFields.Count > 0)
            {
                result.AppendLine($"列欄位數: {pivotTable.ColumnFields.Count}");
            }
            
            if (pivotTable.DataFields != null && pivotTable.DataFields.Count > 0)
            {
                result.AppendLine($"數據欄位數: {pivotTable.DataFields.Count}");
            }
            
            result.AppendLine();
        }

        return await Task.FromResult(result.ToString());
    }
}

