using System.Text.Json.Nodes;
using Aspose.Cells;

namespace AsposeMcpServer.Tools;

public class ExcelSetColumnWidthTool : IAsposeTool
{
    public string Description => "Set column width in an Excel worksheet";

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
            columnIndex = new
            {
                type = "number",
                description = "Column index (0-based)"
            },
            width = new
            {
                type = "number",
                description = "Column width in characters (Excel unit)"
            }
        },
        required = new[] { "path", "columnIndex", "width" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var sheetIndex = arguments?["sheetIndex"]?.GetValue<int>() ?? 0;
        var columnIndex = arguments?["columnIndex"]?.GetValue<int>() ?? throw new ArgumentException("columnIndex is required");
        var width = arguments?["width"]?.GetValue<double>() ?? throw new ArgumentException("width is required");

        using var workbook = new Workbook(path);
        var worksheet = workbook.Worksheets[sheetIndex];

        worksheet.Cells.SetColumnWidth(columnIndex, width);
        workbook.Save(path);

        return await Task.FromResult($"第 {columnIndex} 列寬度已設定為 {width} 字符: {path}");
    }
}

