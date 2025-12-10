using System.Text.Json.Nodes;
using System.Text;
using Aspose.Cells;

namespace AsposeMcpServer.Tools;

public class ExcelGetSheetInfoTool : IAsposeTool
{
    public string Description => "Get detailed information about worksheets in an Excel workbook";

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
                description = "Sheet index to get info for (0-based, optional, if not provided returns all sheets)"
            }
        },
        required = new[] { "path" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var sheetIndex = arguments?["sheetIndex"]?.GetValue<int?>();

        using var workbook = new Workbook(path);
        var result = new StringBuilder();

        result.AppendLine("=== Excel 工作簿資訊 ===\n");
        result.AppendLine($"總工作表數: {workbook.Worksheets.Count}\n");

        if (sheetIndex.HasValue)
        {
            if (sheetIndex.Value < 0 || sheetIndex.Value >= workbook.Worksheets.Count)
            {
                throw new ArgumentException($"工作表索引 {sheetIndex.Value} 超出範圍 (共有 {workbook.Worksheets.Count} 個工作表)");
            }

            var worksheet = workbook.Worksheets[sheetIndex.Value];
            AppendSheetInfo(result, worksheet, sheetIndex.Value);
        }
        else
        {
            for (int i = 0; i < workbook.Worksheets.Count; i++)
            {
                AppendSheetInfo(result, workbook.Worksheets[i], i);
                if (i < workbook.Worksheets.Count - 1) result.AppendLine();
            }
        }

        return await Task.FromResult(result.ToString());
    }

    private void AppendSheetInfo(StringBuilder result, Worksheet worksheet, int index)
    {
        result.AppendLine($"【工作表 {index}: {worksheet.Name}】");
        result.AppendLine($"  可見性: {worksheet.VisibilityType}");
        result.AppendLine($"  最大行: {worksheet.Cells.MaxDataRow + 1}");
        result.AppendLine($"  最大列: {worksheet.Cells.MaxDataColumn + 1}");
        result.AppendLine($"  已使用範圍: {worksheet.Cells.MaxRow + 1} 行 × {worksheet.Cells.MaxColumn + 1} 列");
        result.AppendLine($"  頁面方向: {worksheet.PageSetup.Orientation}");
        result.AppendLine($"  紙張大小: {worksheet.PageSetup.PaperSize}");
        result.AppendLine($"  凍結窗格: 行 {worksheet.FirstVisibleRow}, 列 {worksheet.FirstVisibleColumn}");
    }
}

