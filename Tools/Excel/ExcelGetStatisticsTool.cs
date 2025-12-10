using System.Text.Json.Nodes;
using System.Text;
using Aspose.Cells;

namespace AsposeMcpServer.Tools;

public class ExcelGetStatisticsTool : IAsposeTool
{
    public string Description => "Get comprehensive statistics about an Excel workbook";

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
                description = "Sheet index (0-based, optional, if not provided returns statistics for all sheets)"
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

        result.AppendLine("=== Excel 工作簿統計資訊 ===\n");

        // Workbook level statistics
        result.AppendLine("【工作簿資訊】");
        result.AppendLine($"總工作表數: {workbook.Worksheets.Count}");
        result.AppendLine($"檔案格式: {workbook.FileFormat}");
        result.AppendLine();

        if (sheetIndex.HasValue)
        {
            if (sheetIndex.Value < 0 || sheetIndex.Value >= workbook.Worksheets.Count)
            {
                throw new ArgumentException($"工作表索引 {sheetIndex.Value} 超出範圍 (共有 {workbook.Worksheets.Count} 個工作表)");
            }
            AppendSheetStatistics(result, workbook.Worksheets[sheetIndex.Value], sheetIndex.Value);
        }
        else
        {
            for (int i = 0; i < workbook.Worksheets.Count; i++)
            {
                AppendSheetStatistics(result, workbook.Worksheets[i], i);
                if (i < workbook.Worksheets.Count - 1) result.AppendLine();
            }
        }

        return await Task.FromResult(result.ToString());
    }

    private void AppendSheetStatistics(StringBuilder result, Worksheet worksheet, int index)
    {
        result.AppendLine($"【工作表 {index}: {worksheet.Name}】");
        
        // Basic info
        result.AppendLine("基本資訊:");
        result.AppendLine($"  可見性: {worksheet.VisibilityType}");
        result.AppendLine($"  最大數據行: {worksheet.Cells.MaxDataRow + 1}");
        result.AppendLine($"  最大數據列: {worksheet.Cells.MaxDataColumn + 1}");
        result.AppendLine($"  已使用範圍: {worksheet.Cells.MaxRow + 1} 行 × {worksheet.Cells.MaxColumn + 1} 列");
        result.AppendLine();

        // Cell statistics
        int cellCount = 0;
        int formulaCount = 0;
        int emptyCount = 0;
        int textCount = 0;
        int numberCount = 0;
        int dateCount = 0;

        var maxRow = Math.Min(worksheet.Cells.MaxDataRow + 1, 10000); // Limit for performance
        var maxCol = Math.Min(worksheet.Cells.MaxDataColumn + 1, 1000);

        for (int row = 0; row < maxRow; row++)
        {
            for (int col = 0; col < maxCol; col++)
            {
                var cell = worksheet.Cells[row, col];
                if (cell.Value != null)
                {
                    cellCount++;
                    var value = cell.Value;
                    if (cell.Formula != null && cell.Formula.Length > 0)
                    {
                        formulaCount++;
                    }
                    else if (value is string str && string.IsNullOrWhiteSpace(str))
                    {
                        emptyCount++;
                    }
                    else if (value is string)
                    {
                        textCount++;
                    }
                    else if (value is double || value is int || value is decimal)
                    {
                        numberCount++;
                    }
                    else if (value is DateTime)
                    {
                        dateCount++;
                    }
                }
            }
        }

        result.AppendLine("單元格統計:");
        result.AppendLine($"  有數據的單元格: {cellCount}");
        result.AppendLine($"  包含公式: {formulaCount}");
        result.AppendLine($"  文字: {textCount}");
        result.AppendLine($"  數字: {numberCount}");
        result.AppendLine($"  日期: {dateCount}");
        result.AppendLine();

        // Charts
        result.AppendLine($"圖表數: {worksheet.Charts.Count}");
        
        // Pivot tables
        result.AppendLine($"樞紐表數: {worksheet.PivotTables.Count}");
        
        // Conditional formatting
        result.AppendLine($"條件格式數: {worksheet.ConditionalFormattings.Count}");
        
        // Data validation
        result.AppendLine($"數據驗證數: {worksheet.Validations.Count}");
        
        // Hyperlinks
        result.AppendLine($"超連結數: {worksheet.Hyperlinks.Count}");
        
        // Pictures
        result.AppendLine($"圖片數: {worksheet.Pictures.Count}");
        
        // Comments
        result.AppendLine($"註解數: {worksheet.Comments.Count}");
        
        // Protection
        result.AppendLine($"保護狀態: {(worksheet.Protection.IsProtectedWithPassword ? "已保護" : "未保護")}");
        
        // Page setup
        result.AppendLine($"頁面方向: {worksheet.PageSetup.Orientation}");
        result.AppendLine($"紙張大小: {worksheet.PageSetup.PaperSize}");
        result.AppendLine($"凍結窗格: 行 {worksheet.FirstVisibleRow}, 列 {worksheet.FirstVisibleColumn}");
    }
}

