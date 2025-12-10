using System.Text.Json.Nodes;
using Aspose.Cells;

namespace AsposeMcpServer.Tools;

public class ExcelClearCellTool : IAsposeTool
{
    public string Description => "Clear cell content, format, or both in Excel";

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
            cell = new
            {
                type = "string",
                description = "Cell reference (e.g., 'A1')"
            },
            clearContent = new
            {
                type = "boolean",
                description = "Clear cell content (optional, default: true)"
            },
            clearFormat = new
            {
                type = "boolean",
                description = "Clear cell format (optional, default: false)"
            }
        },
        required = new[] { "path", "cell" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var sheetIndex = arguments?["sheetIndex"]?.GetValue<int>() ?? 0;
        var cell = arguments?["cell"]?.GetValue<string>() ?? throw new ArgumentException("cell is required");
        var clearContent = arguments?["clearContent"]?.GetValue<bool?>() ?? true;
        var clearFormat = arguments?["clearFormat"]?.GetValue<bool?>() ?? false;

        using var workbook = new Workbook(path);
        if (sheetIndex < 0 || sheetIndex >= workbook.Worksheets.Count)
        {
            throw new ArgumentException($"sheetIndex must be between 0 and {workbook.Worksheets.Count - 1}");
        }

        var worksheet = workbook.Worksheets[sheetIndex];
        var cellObj = worksheet.Cells[cell];

        if (clearContent && clearFormat)
        {
            cellObj.PutValue("");
            var defaultStyle = workbook.CreateStyle();
            cellObj.SetStyle(defaultStyle);
        }
        else if (clearContent)
        {
            cellObj.PutValue("");
        }
        else if (clearFormat)
        {
            var defaultStyle = workbook.CreateStyle();
            cellObj.SetStyle(defaultStyle);
        }

        workbook.Save(path);
        return await Task.FromResult($"Cell {cell} cleared in sheet {sheetIndex}: {path}");
    }
}

