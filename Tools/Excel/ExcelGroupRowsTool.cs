using System.Text.Json.Nodes;
using Aspose.Cells;

namespace AsposeMcpServer.Tools;

public class ExcelGroupRowsTool : IAsposeTool
{
    public string Description => "Group rows in Excel (create outline group)";

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
            startRow = new
            {
                type = "number",
                description = "Start row index (0-based)"
            },
            endRow = new
            {
                type = "number",
                description = "End row index (0-based)"
            },
            isCollapsed = new
            {
                type = "boolean",
                description = "Collapse group initially (optional, default: false)"
            }
        },
        required = new[] { "path", "startRow", "endRow" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var sheetIndex = arguments?["sheetIndex"]?.GetValue<int>() ?? 0;
        var startRow = arguments?["startRow"]?.GetValue<int>() ?? throw new ArgumentException("startRow is required");
        var endRow = arguments?["endRow"]?.GetValue<int>() ?? throw new ArgumentException("endRow is required");
        var isCollapsed = arguments?["isCollapsed"]?.GetValue<bool?>() ?? false;

        using var workbook = new Workbook(path);
        if (sheetIndex < 0 || sheetIndex >= workbook.Worksheets.Count)
        {
            throw new ArgumentException($"sheetIndex must be between 0 and {workbook.Worksheets.Count - 1}");
        }

        var worksheet = workbook.Worksheets[sheetIndex];
        worksheet.Cells.GroupRows(startRow, endRow, isCollapsed);

        workbook.Save(path);
        return await Task.FromResult($"Rows {startRow}-{endRow} grouped in sheet {sheetIndex}: {path}");
    }
}

