using System.Text.Json.Nodes;
using Aspose.Cells;

namespace AsposeMcpServer.Tools;

public class ExcelFreezePanesTool : IAsposeTool
{
    public string Description => "Freeze or unfreeze panes in an Excel worksheet";

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
            row = new
            {
                type = "number",
                description = "Row index to freeze at (0-based, optional, 0 to unfreeze)"
            },
            column = new
            {
                type = "number",
                description = "Column index to freeze at (0-based, optional, 0 to unfreeze)"
            }
        },
        required = new[] { "path" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var sheetIndex = arguments?["sheetIndex"]?.GetValue<int>() ?? 0;
        var row = arguments?["row"]?.GetValue<int>() ?? 0;
        var column = arguments?["column"]?.GetValue<int>() ?? 0;

        using var workbook = new Workbook(path);
        var worksheet = workbook.Worksheets[sheetIndex];

        if (row == 0 && column == 0)
        {
            worksheet.FreezePanes(0, 0, 0, 0);
            workbook.Save(path);
            return await Task.FromResult($"已取消凍結窗格: {path}");
        }
        else
        {
            worksheet.FreezePanes(row, column, row, column);
            workbook.Save(path);
            return await Task.FromResult($"已凍結窗格 (行 {row}, 列 {column}): {path}");
        }
    }
}

