using System.Text.Json.Nodes;
using Aspose.Cells;

namespace AsposeMcpServer.Tools;

public class ExcelDeleteDataValidationTool : IAsposeTool
{
    public string Description => "Delete data validation from an Excel worksheet";

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
            validationIndex = new
            {
                type = "number",
                description = "Data validation index to delete (0-based)"
            }
        },
        required = new[] { "path", "validationIndex" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var sheetIndex = arguments?["sheetIndex"]?.GetValue<int>() ?? 0;
        var validationIndex = arguments?["validationIndex"]?.GetValue<int>() ?? throw new ArgumentException("validationIndex is required");

        using var workbook = new Workbook(path);
        
        if (sheetIndex < 0 || sheetIndex >= workbook.Worksheets.Count)
        {
            throw new ArgumentException($"工作表索引 {sheetIndex} 超出範圍 (共有 {workbook.Worksheets.Count} 個工作表)");
        }

        var worksheet = workbook.Worksheets[sheetIndex];
        var validations = worksheet.Validations;
        
        if (validationIndex < 0 || validationIndex >= validations.Count)
        {
            throw new ArgumentException($"數據驗證索引 {validationIndex} 超出範圍 (工作表共有 {validations.Count} 個數據驗證)");
        }

        validations.RemoveAt(validationIndex);
        workbook.Save(path);
        
        var remainingCount = validations.Count;
        
        return await Task.FromResult($"成功刪除數據驗證 #{validationIndex}\n工作表剩餘數據驗證數: {remainingCount}\n輸出: {path}");
    }
}

