using System.Text.Json.Nodes;
using System.Text;
using Aspose.Cells;

namespace AsposeMcpServer.Tools;

public class ExcelGetProtectionTool : IAsposeTool
{
    public string Description => "Get protection settings for Excel workbook and worksheets";

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
                description = "Sheet index (0-based, optional, if not provided returns protection for all sheets)"
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

        result.AppendLine("=== Excel 保護設定資訊 ===\n");

        // Workbook protection - Aspose.Cells doesn't expose workbook protection status directly
        // We can only check worksheet protection
        result.AppendLine("【工作簿保護】");
        result.AppendLine("注意: 工作簿保護狀態需要通過保護方法檢查");
        result.AppendLine();

        if (sheetIndex.HasValue)
        {
            if (sheetIndex.Value < 0 || sheetIndex.Value >= workbook.Worksheets.Count)
            {
                throw new ArgumentException($"工作表索引 {sheetIndex.Value} 超出範圍 (共有 {workbook.Worksheets.Count} 個工作表)");
            }
            AppendSheetProtection(result, workbook.Worksheets[sheetIndex.Value], sheetIndex.Value);
        }
        else
        {
            for (int i = 0; i < workbook.Worksheets.Count; i++)
            {
                AppendSheetProtection(result, workbook.Worksheets[i], i);
                if (i < workbook.Worksheets.Count - 1) result.AppendLine();
            }
        }

        return await Task.FromResult(result.ToString());
    }

    private void AppendSheetProtection(StringBuilder result, Worksheet worksheet, int index)
    {
        var protection = worksheet.Protection;
        result.AppendLine($"【工作表 {index}: {worksheet.Name}】");
        result.AppendLine($"保護狀態: {(protection.IsProtectedWithPassword ? "已保護" : "未保護")}");
        result.AppendLine($"允許選擇鎖定單元格: {protection.AllowSelectingLockedCell}");
        result.AppendLine($"允許選擇未鎖定單元格: {protection.AllowSelectingUnlockedCell}");
        result.AppendLine($"允許格式化單元格: {protection.AllowFormattingCell}");
        result.AppendLine($"允許格式化列: {protection.AllowFormattingColumn}");
        result.AppendLine($"允許格式化行: {protection.AllowFormattingRow}");
        result.AppendLine($"允許插入列: {protection.AllowInsertingColumn}");
        result.AppendLine($"允許插入行: {protection.AllowInsertingRow}");
        result.AppendLine($"允許插入超連結: {protection.AllowInsertingHyperlink}");
        result.AppendLine($"允許刪除列: {protection.AllowDeletingColumn}");
        result.AppendLine($"允許刪除行: {protection.AllowDeletingRow}");
        result.AppendLine($"允許排序: {protection.AllowSorting}");
        result.AppendLine($"允許自動篩選: {protection.AllowFiltering}");
        result.AppendLine($"允許使用樞紐表: {protection.AllowUsingPivotTable}");
        result.AppendLine($"允許編輯對象: {protection.AllowEditingObject}");
        result.AppendLine($"允許編輯場景: {protection.AllowEditingScenario}");
    }
}

