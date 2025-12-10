using System.Text.Json.Nodes;
using System.Text;
using Aspose.Cells;

namespace AsposeMcpServer.Tools;

public class ExcelGetFilterStatusTool : IAsposeTool
{
    public string Description => "Get auto filter status from an Excel worksheet";

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
        var autoFilter = worksheet.AutoFilter;
        var result = new StringBuilder();

        result.AppendLine($"=== 工作表 '{worksheet.Name}' 的自動篩選狀態 ===\n");

        if (string.IsNullOrEmpty(autoFilter.Range))
        {
            result.AppendLine("狀態: 未啟用自動篩選");
            return await Task.FromResult(result.ToString());
        }

        result.AppendLine($"狀態: 已啟用");
        result.AppendLine($"篩選範圍: {autoFilter.Range}");
        
        // Get filter columns information
        if (autoFilter.FilterColumns != null && autoFilter.FilterColumns.Count > 0)
        {
            result.AppendLine($"\n篩選欄位數: {autoFilter.FilterColumns.Count}");
            for (int i = 0; i < autoFilter.FilterColumns.Count; i++)
            {
                var filterColumn = autoFilter.FilterColumns[i];
                // Check if filter has criteria set
                bool hasFilter = filterColumn.FilterType != FilterType.None;
                result.AppendLine($"  欄位 {i}: 已篩選={hasFilter}");
            }
        }

        return await Task.FromResult(result.ToString());
    }
}

