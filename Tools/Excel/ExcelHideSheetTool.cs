using System.Text.Json.Nodes;
using Aspose.Cells;

namespace AsposeMcpServer.Tools;

public class ExcelHideSheetTool : IAsposeTool
{
    public string Description => "Hide or show a worksheet in an Excel workbook";

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
                description = "Sheet index to hide/show (0-based)"
            },
            hidden = new
            {
                type = "boolean",
                description = "True to hide, false to show (default: true)"
            }
        },
        required = new[] { "path", "sheetIndex" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var sheetIndex = arguments?["sheetIndex"]?.GetValue<int>() ?? throw new ArgumentException("sheetIndex is required");
        var hidden = arguments?["hidden"]?.GetValue<bool>() ?? true;

        using var workbook = new Workbook(path);

        if (sheetIndex < 0 || sheetIndex >= workbook.Worksheets.Count)
        {
            throw new ArgumentException($"工作表索引 {sheetIndex} 超出範圍 (共有 {workbook.Worksheets.Count} 個工作表)");
        }

        var sheet = workbook.Worksheets[sheetIndex];
        var sheetName = sheet.Name;

        if (hidden)
        {
            sheet.VisibilityType = VisibilityType.Hidden;
            workbook.Save(path);
            return await Task.FromResult($"工作表 '{sheetName}' 已隱藏: {path}");
        }
        else
        {
            sheet.VisibilityType = VisibilityType.Visible;
            workbook.Save(path);
            return await Task.FromResult($"工作表 '{sheetName}' 已顯示: {path}");
        }
    }
}

