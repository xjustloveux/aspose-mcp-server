using System.Text.Json.Nodes;
using Aspose.Cells;

namespace AsposeMcpServer.Tools;

public class ExcelDeleteCellsTool : IAsposeTool
{
    public string Description => "Delete cells (shift cells left or up) in Excel";

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
                description = "Range to delete (e.g., 'A1:C5')"
            },
            shiftDirection = new
            {
                type = "string",
                description = "Shift direction: 'Left' or 'Up'",
                @enum = new[] { "Left", "Up" }
            }
        },
        required = new[] { "path", "range", "shiftDirection" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var sheetIndex = arguments?["sheetIndex"]?.GetValue<int?>() ?? 0;
        var range = arguments?["range"]?.GetValue<string>() ?? throw new ArgumentException("range is required");
        var shiftDirection = arguments?["shiftDirection"]?.GetValue<string>() ?? throw new ArgumentException("shiftDirection is required");

        using var workbook = new Workbook(path);
        if (sheetIndex < 0 || sheetIndex >= workbook.Worksheets.Count)
        {
            throw new ArgumentException($"sheetIndex must be between 0 and {workbook.Worksheets.Count - 1}");
        }

        var worksheet = workbook.Worksheets[sheetIndex];
        var rangeObj = worksheet.Cells.CreateRange(range);

        var shiftType = shiftDirection.ToLower() == "left" ? ShiftType.Left : ShiftType.Up;
        worksheet.Cells.DeleteRange(rangeObj.FirstRow, rangeObj.FirstColumn, rangeObj.RowCount, rangeObj.ColumnCount, shiftType);

        workbook.Save(path);
        return await Task.FromResult($"Cells deleted in range {range}, shifted {shiftDirection}: {path}");
    }
}

