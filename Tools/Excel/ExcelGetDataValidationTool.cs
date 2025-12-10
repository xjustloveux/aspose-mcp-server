using System.Text.Json.Nodes;
using System.Text;
using Aspose.Cells;

namespace AsposeMcpServer.Tools;

public class ExcelGetDataValidationTool : IAsposeTool
{
    public string Description => "Get all data validation information from an Excel worksheet";

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
        var validations = worksheet.Validations;
        var result = new StringBuilder();

        result.AppendLine($"=== 工作表 '{worksheet.Name}' 的數據驗證資訊 ===\n");
        result.AppendLine($"總數據驗證數: {validations.Count}\n");

        if (validations.Count == 0)
        {
            result.AppendLine("未找到數據驗證");
            return await Task.FromResult(result.ToString());
        }

        for (int i = 0; i < validations.Count; i++)
        {
            var validation = validations[i];
            result.AppendLine($"【數據驗證 {i}】");
            result.AppendLine($"類型: {validation.Type}");
            result.AppendLine($"運算符: {validation.Operator}");
            
            // Note: Validation area information may not be directly accessible
            // The validation is applied but area details require different API access
            result.AppendLine("應用範圍: 已應用（詳細範圍資訊需通過其他方式獲取）");
            
            if (!string.IsNullOrEmpty(validation.Formula1))
            {
                result.AppendLine($"公式1: {validation.Formula1}");
            }
            if (!string.IsNullOrEmpty(validation.Formula2))
            {
                result.AppendLine($"公式2: {validation.Formula2}");
            }
            if (!string.IsNullOrEmpty(validation.ErrorMessage))
            {
                result.AppendLine($"錯誤訊息: {validation.ErrorMessage}");
            }
            if (!string.IsNullOrEmpty(validation.InputMessage))
            {
                result.AppendLine($"輸入訊息: {validation.InputMessage}");
            }
            result.AppendLine($"顯示錯誤: {validation.ShowError}");
            result.AppendLine($"顯示輸入: {validation.ShowInput}");
            result.AppendLine($"下拉列表: {validation.InCellDropDown}");
            result.AppendLine();
        }

        return await Task.FromResult(result.ToString());
    }
}

