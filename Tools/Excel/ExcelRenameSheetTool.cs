using System.Text.Json.Nodes;
using Aspose.Cells;
using AsposeMcpServer.Core;

namespace AsposeMcpServer.Tools;

public class ExcelRenameSheetTool : IAsposeTool
{
    public string Description => "Rename a worksheet in an Excel workbook";

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
                description = "Sheet index to rename (0-based)"
            },
            newName = new
            {
                type = "string",
                description = "New name for the worksheet"
            }
        },
        required = new[] { "path", "sheetIndex", "newName" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var sheetIndex = arguments?["sheetIndex"]?.GetValue<int>() ?? throw new ArgumentException("sheetIndex is required");
        var newName = arguments?["newName"]?.GetValue<string>() ?? throw new ArgumentException("newName is required");

        using var workbook = new Workbook(path);

        if (sheetIndex < 0 || sheetIndex >= workbook.Worksheets.Count)
        {
            throw new ArgumentException($"工作表索引 {sheetIndex} 超出範圍 (共有 {workbook.Worksheets.Count} 個工作表)");
        }

        var oldName = workbook.Worksheets[sheetIndex].Name;
        workbook.Worksheets[sheetIndex].Name = newName;
        workbook.Save(path);

        return await Task.FromResult($"工作表 '{oldName}' 已重命名為 '{newName}': {path}");
    }
}

