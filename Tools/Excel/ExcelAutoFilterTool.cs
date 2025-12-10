using System.Text.Json.Nodes;
using Aspose.Cells;

namespace AsposeMcpServer.Tools;

public class ExcelAutoFilterTool : IAsposeTool
{
    public string Description => "Apply or remove auto filter on a range in an Excel worksheet";

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
            range = new
            {
                type = "string",
                description = "Cell range to apply filter (e.g., 'A1:C10', should include header row)"
            },
            enable = new
            {
                type = "boolean",
                description = "True to enable filter, false to remove (default: true)"
            }
        },
        required = new[] { "path", "range" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var sheetIndex = arguments?["sheetIndex"]?.GetValue<int>() ?? 0;
        var range = arguments?["range"]?.GetValue<string>() ?? throw new ArgumentException("range is required");
        var enable = arguments?["enable"]?.GetValue<bool>() ?? true;

        using var workbook = new Workbook(path);
        var worksheet = workbook.Worksheets[sheetIndex];
        var cells = worksheet.Cells;
        var cellRange = cells.CreateRange(range);

        if (enable)
        {
            worksheet.AutoFilter.Range = cellRange.Address;
            workbook.Save(path);
            return await Task.FromResult($"範圍 {range} 已啟用自動篩選: {path}");
        }
        else
        {
            worksheet.AutoFilter.Range = "";
            workbook.Save(path);
            return await Task.FromResult($"已移除自動篩選: {path}");
        }
    }
}

