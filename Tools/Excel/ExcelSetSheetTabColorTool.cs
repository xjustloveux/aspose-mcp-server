using System.Text.Json.Nodes;
using Aspose.Cells;
using System.Drawing;

namespace AsposeMcpServer.Tools;

public class ExcelSetSheetTabColorTool : IAsposeTool
{
    public string Description => "Set worksheet tab color in Excel";

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
            color = new
            {
                type = "string",
                description = "Color in hex format (e.g., 'FF0000' for red, or color name like 'Red')"
            }
        },
        required = new[] { "path", "color" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var sheetIndex = arguments?["sheetIndex"]?.GetValue<int?>() ?? 0;
        var colorStr = arguments?["color"]?.GetValue<string>() ?? throw new ArgumentException("color is required");

        using var workbook = new Workbook(path);
        if (sheetIndex < 0 || sheetIndex >= workbook.Worksheets.Count)
        {
            throw new ArgumentException($"sheetIndex must be between 0 and {workbook.Worksheets.Count - 1}");
        }

        var worksheet = workbook.Worksheets[sheetIndex];

        Color color;
        if (colorStr.StartsWith("#"))
        {
            colorStr = colorStr.Substring(1);
        }

        if (colorStr.Length == 6)
        {
            var r = Convert.ToInt32(colorStr.Substring(0, 2), 16);
            var g = Convert.ToInt32(colorStr.Substring(2, 2), 16);
            var b = Convert.ToInt32(colorStr.Substring(4, 2), 16);
            color = Color.FromArgb(r, g, b);
        }
        else
        {
            color = Color.FromName(colorStr);
        }

        worksheet.TabColor = color;

        workbook.Save(path);
        return await Task.FromResult($"Sheet tab color set to {colorStr}: {path}");
    }
}

