using System.Text.Json.Nodes;
using Aspose.Cells;

namespace AsposeMcpServer.Tools;

public class ExcelDeleteSheetTool : IAsposeTool
{
    public string Description => "Delete a worksheet from an Excel workbook";

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
                description = "Sheet index to delete (0-based)"
            }
        },
        required = new[] { "path", "sheetIndex" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var sheetIndex = arguments?["sheetIndex"]?.GetValue<int>() ?? throw new ArgumentException("sheetIndex is required");

        using var workbook = new Workbook(path);

        if (sheetIndex < 0 || sheetIndex >= workbook.Worksheets.Count)
        {
            throw new ArgumentException($"工作表索引 {sheetIndex} 超出範圍 (共有 {workbook.Worksheets.Count} 個工作表)");
        }

        if (workbook.Worksheets.Count <= 1)
        {
            throw new InvalidOperationException("無法刪除最後一個工作表");
        }

        var sheetName = workbook.Worksheets[sheetIndex].Name;
        workbook.Worksheets.RemoveAt(sheetIndex);
        workbook.Save(path);

        return await Task.FromResult($"工作表 '{sheetName}' (索引 {sheetIndex}) 已刪除: {path}");
    }
}

