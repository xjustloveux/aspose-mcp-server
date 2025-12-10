using System.Text.Json.Nodes;
using Aspose.Cells;

namespace AsposeMcpServer.Tools;

public class ExcelSetCellLockedTool : IAsposeTool
{
    public string Description => "Set cell locked/unlocked status in Excel (for protection)";

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
                description = "Cell or range (e.g., 'A1' or 'A1:C5')"
            },
            locked = new
            {
                type = "boolean",
                description = "Locked status (true = locked, false = unlocked)"
            }
        },
        required = new[] { "path", "range", "locked" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var sheetIndex = arguments?["sheetIndex"]?.GetValue<int?>() ?? 0;
        var range = arguments?["range"]?.GetValue<string>() ?? throw new ArgumentException("range is required");
        var locked = arguments?["locked"]?.GetValue<bool>() ?? throw new ArgumentException("locked is required");

        using var workbook = new Workbook(path);
        if (sheetIndex < 0 || sheetIndex >= workbook.Worksheets.Count)
        {
            throw new ArgumentException($"sheetIndex must be between 0 and {workbook.Worksheets.Count - 1}");
        }

        var worksheet = workbook.Worksheets[sheetIndex];
        var rangeObj = worksheet.Cells.CreateRange(range);

        // Set locked status for all cells in range
        var style = workbook.CreateStyle();
        style.IsLocked = locked;
        rangeObj.ApplyStyle(style, new StyleFlag { Locked = true });

        workbook.Save(path);
        return await Task.FromResult($"Cell(s) in range {range} set to {(locked ? "locked" : "unlocked")}: {path}");
    }
}

