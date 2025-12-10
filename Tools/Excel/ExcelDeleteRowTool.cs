using System.Text.Json.Nodes;
using Aspose.Cells;

namespace AsposeMcpServer.Tools;

public class ExcelDeleteRowTool : IAsposeTool
{
    public string Description => "Delete one or more rows from an Excel worksheet";

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
                description = "Row index to delete (0-based)"
            },
            count = new
            {
                type = "number",
                description = "Number of rows to delete (default: 1)"
            }
        },
        required = new[] { "path", "rowIndex" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var sheetIndex = arguments?["sheetIndex"]?.GetValue<int>() ?? 0;
        var rowIndex = arguments?["rowIndex"]?.GetValue<int>() ?? throw new ArgumentException("rowIndex is required");
        var count = arguments?["count"]?.GetValue<int>() ?? 1;

        using var workbook = new Workbook(path);
        var worksheet = workbook.Worksheets[sheetIndex];

        for (int i = 0; i < count; i++)
        {
            worksheet.Cells.DeleteRow(rowIndex);
        }
        workbook.Save(path);

        return await Task.FromResult($"已刪除第 {rowIndex} 行起的 {count} 行: {path}");
    }
}

