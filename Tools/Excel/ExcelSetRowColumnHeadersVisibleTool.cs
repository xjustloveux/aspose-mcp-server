using System.Text.Json.Nodes;
using Aspose.Cells;

namespace AsposeMcpServer.Tools;

public class ExcelSetRowColumnHeadersVisibleTool : IAsposeTool
{
    public string Description => "Set row and column headers visibility in Excel worksheet";

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
            visible = new
            {
                type = "boolean",
                description = "Headers visibility (true = visible, false = hidden)"
            }
        },
        required = new[] { "path", "visible" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var sheetIndex = arguments?["sheetIndex"]?.GetValue<int?>() ?? 0;
        var visible = arguments?["visible"]?.GetValue<bool>() ?? throw new ArgumentException("visible is required");

        using var workbook = new Workbook(path);
        if (sheetIndex < 0 || sheetIndex >= workbook.Worksheets.Count)
        {
            throw new ArgumentException($"sheetIndex must be between 0 and {workbook.Worksheets.Count - 1}");
        }

        var worksheet = workbook.Worksheets[sheetIndex];
        worksheet.IsRowColumnHeadersVisible = visible;

        workbook.Save(path);
        return await Task.FromResult($"RowColumnHeaders visibility set to {(visible ? "visible" : "hidden")}: {path}");
    }
}

