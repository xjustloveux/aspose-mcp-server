using System.Text.Json.Nodes;
using Aspose.Cells;

namespace AsposeMcpServer.Tools;

public class ExcelUngroupColumnsTool : IAsposeTool
{
    public string Description => "Ungroup columns in Excel (remove outline group)";

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
            startColumn = new
            {
                type = "number",
                description = "Start column index (0-based)"
            },
            endColumn = new
            {
                type = "number",
                description = "End column index (0-based)"
            }
        },
        required = new[] { "path", "startColumn", "endColumn" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var sheetIndex = arguments?["sheetIndex"]?.GetValue<int>() ?? 0;
        var startColumn = arguments?["startColumn"]?.GetValue<int>() ?? throw new ArgumentException("startColumn is required");
        var endColumn = arguments?["endColumn"]?.GetValue<int>() ?? throw new ArgumentException("endColumn is required");

        using var workbook = new Workbook(path);
        if (sheetIndex < 0 || sheetIndex >= workbook.Worksheets.Count)
        {
            throw new ArgumentException($"sheetIndex must be between 0 and {workbook.Worksheets.Count - 1}");
        }

        var worksheet = workbook.Worksheets[sheetIndex];
        worksheet.Cells.UngroupColumns(startColumn, endColumn);

        workbook.Save(path);
        return await Task.FromResult($"Columns {startColumn}-{endColumn} ungrouped in sheet {sheetIndex}: {path}");
    }
}

