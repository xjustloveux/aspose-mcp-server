using System.Text.Json.Nodes;
using Aspose.Cells;

namespace AsposeMcpServer.Tools;

public class ExcelSetRowHeightTool : IAsposeTool
{
    public string Description => "Set row height in an Excel worksheet";

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
            rowIndex = new
            {
                type = "number",
                description = "Row index (0-based)"
            },
            height = new
            {
                type = "number",
                description = "Row height in points"
            }
        },
        required = new[] { "path", "rowIndex", "height" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var sheetIndex = arguments?["sheetIndex"]?.GetValue<int>() ?? 0;
        var rowIndex = arguments?["rowIndex"]?.GetValue<int>() ?? throw new ArgumentException("rowIndex is required");
        var height = arguments?["height"]?.GetValue<double>() ?? throw new ArgumentException("height is required");

        using var workbook = new Workbook(path);
        var worksheet = workbook.Worksheets[sheetIndex];

        worksheet.Cells.SetRowHeight(rowIndex, height);
        workbook.Save(path);

        return await Task.FromResult($"第 {rowIndex} 行高度已設定為 {height} 點: {path}");
    }
}

