using System.Text.Json.Nodes;
using System.Text;
using Aspose.Cells;

namespace AsposeMcpServer.Tools;

public class ExcelGetFreezePanesTool : IAsposeTool
{
    public string Description => "Get freeze panes status from an Excel worksheet";

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
        var result = new StringBuilder();

        result.AppendLine($"=== 工作表 '{worksheet.Name}' 的凍結窗格狀態 ===\n");

        var firstVisibleRow = worksheet.FirstVisibleRow;
        var firstVisibleColumn = worksheet.FirstVisibleColumn;

        if (firstVisibleRow == 0 && firstVisibleColumn == 0)
        {
            result.AppendLine("狀態: 未凍結窗格");
        }
        else
        {
            result.AppendLine("狀態: 已凍結窗格");
            result.AppendLine($"凍結行: {firstVisibleRow}");
            result.AppendLine($"凍結列: {firstVisibleColumn}");
            result.AppendLine($"凍結位置: 行 {firstVisibleRow + 1} 和列 {firstVisibleColumn + 1} 之前");
        }

        return await Task.FromResult(result.ToString());
    }
}

