using System.Text.Json.Nodes;
using Aspose.Cells;

namespace AsposeMcpServer.Tools;

public class ExcelInsertCellsTool : IAsposeTool
{
    public string Description => "Insert cells (shift cells right or down) in Excel";

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
                description = "Range where to insert cells (e.g., 'A1:C5')"
            },
            shiftDirection = new
            {
                type = "string",
                description = "Shift direction: 'Right' or 'Down'",
                @enum = new[] { "Right", "Down" }
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

        var shiftType = shiftDirection.ToLower() == "right" ? ShiftType.Right : ShiftType.Down;
        
        // Insert cells by inserting rows/columns
        if (shiftType == ShiftType.Down)
        {
            // Insert rows
            for (int i = 0; i < rangeObj.RowCount; i++)
            {
                worksheet.Cells.InsertRow(rangeObj.FirstRow);
            }
        }
        else
        {
            // Insert columns
            for (int i = 0; i < rangeObj.ColumnCount; i++)
            {
                worksheet.Cells.InsertColumn(rangeObj.FirstColumn);
            }
        }

        workbook.Save(path);
        return await Task.FromResult($"Cells inserted in range {range}, shifted {shiftDirection}: {path}");
    }
}

