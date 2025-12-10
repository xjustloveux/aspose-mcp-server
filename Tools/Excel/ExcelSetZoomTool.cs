using System.Text.Json.Nodes;
using Aspose.Cells;

namespace AsposeMcpServer.Tools;

public class ExcelSetZoomTool : IAsposeTool
{
    public string Description => "Set worksheet zoom level in Excel";

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
            zoom = new
            {
                type = "number",
                description = "Zoom level (10-400, where 100 = 100%)"
            }
        },
        required = new[] { "path", "zoom" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var sheetIndex = arguments?["sheetIndex"]?.GetValue<int?>() ?? 0;
        var zoom = arguments?["zoom"]?.GetValue<int>() ?? throw new ArgumentException("zoom is required");

        if (zoom < 10 || zoom > 400)
        {
            throw new ArgumentException("Zoom must be between 10 and 400");
        }

        using var workbook = new Workbook(path);
        if (sheetIndex < 0 || sheetIndex >= workbook.Worksheets.Count)
        {
            throw new ArgumentException($"sheetIndex must be between 0 and {workbook.Worksheets.Count - 1}");
        }

        var worksheet = workbook.Worksheets[sheetIndex];
        worksheet.Zoom = zoom;

        workbook.Save(path);
        return await Task.FromResult($"Zoom level set to {zoom}% for sheet {sheetIndex}: {path}");
    }
}

