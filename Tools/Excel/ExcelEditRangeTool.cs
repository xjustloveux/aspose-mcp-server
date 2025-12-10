using System.Text.Json.Nodes;
using Aspose.Cells;

namespace AsposeMcpServer.Tools;

public class ExcelEditRangeTool : IAsposeTool
{
    public string Description => "Edit range of cells with data array in Excel";

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
                description = "Cell range (e.g., 'A1:C5')"
            },
            data = new
            {
                type = "array",
                description = "2D array of cell data",
                items = new
                {
                    type = "array",
                    items = new { type = "string" }
                }
            },
            clearRange = new
            {
                type = "boolean",
                description = "Clear range before writing (optional, default: false)"
            }
        },
        required = new[] { "path", "range", "data" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var sheetIndex = arguments?["sheetIndex"]?.GetValue<int>() ?? 0;
        var range = arguments?["range"]?.GetValue<string>() ?? throw new ArgumentException("range is required");
        var dataArray = arguments?["data"]?.AsArray() ?? throw new ArgumentException("data is required");
        var clearRange = arguments?["clearRange"]?.GetValue<bool?>() ?? false;

        using var workbook = new Workbook(path);
        if (sheetIndex < 0 || sheetIndex >= workbook.Worksheets.Count)
        {
            throw new ArgumentException($"sheetIndex must be between 0 and {workbook.Worksheets.Count - 1}");
        }

        var worksheet = workbook.Worksheets[sheetIndex];
        var cells = worksheet.Cells;
        var cellRange = cells.CreateRange(range);

        if (clearRange)
        {
            // Clear content by iterating through cells
            for (int i = cellRange.FirstRow; i <= cellRange.FirstRow + cellRange.RowCount - 1; i++)
            {
                for (int j = cellRange.FirstColumn; j <= cellRange.FirstColumn + cellRange.ColumnCount - 1; j++)
                {
                    cells[i, j].PutValue("");
                }
            }
        }

        var startRow = cellRange.FirstRow;
        var startCol = cellRange.FirstColumn;

        for (int i = 0; i < dataArray.Count; i++)
        {
            var rowArray = dataArray[i]?.AsArray();
            if (rowArray != null)
            {
                for (int j = 0; j < rowArray.Count; j++)
                {
                    var value = rowArray[j]?.GetValue<string>();
                    cells[startRow + i, startCol + j].PutValue(value ?? "");
                }
            }
        }

        workbook.Save(path);
        return await Task.FromResult($"Range {range} edited in sheet {sheetIndex}: {path}");
    }
}

