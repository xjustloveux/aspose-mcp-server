using System.Text.Json.Nodes;
using Aspose.Cells;

namespace AsposeMcpServer.Tools;

public class ExcelWriteRangeTool : IAsposeTool
{
    public string Description => "Write data to a range of cells in Excel";

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
            startCell = new
            {
                type = "string",
                description = "Starting cell (e.g., 'A1')"
            },
            data = new
            {
                type = "array",
                description = "2D array of data to write",
                items = new
                {
                    type = "array",
                    items = new { type = "string" }
                }
            }
        },
        required = new[] { "path", "startCell", "data" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var sheetIndex = arguments?["sheetIndex"]?.GetValue<int>() ?? 0;
        var startCell = arguments?["startCell"]?.GetValue<string>() ?? throw new ArgumentException("startCell is required");
        var dataArray = arguments?["data"]?.AsArray() ?? throw new ArgumentException("data is required");

        using var workbook = new Workbook(path);
        if (sheetIndex < 0 || sheetIndex >= workbook.Worksheets.Count)
        {
            throw new ArgumentException($"sheetIndex must be between 0 and {workbook.Worksheets.Count - 1}");
        }

        var worksheet = workbook.Worksheets[sheetIndex];

        var startCellObj = worksheet.Cells[startCell];
        int startRow = startCellObj.Row;
        int startCol = startCellObj.Column;

        for (int i = 0; i < dataArray.Count; i++)
        {
            var rowArray = dataArray[i]?.AsArray();
            if (rowArray != null)
            {
                for (int j = 0; j < rowArray.Count; j++)
                {
                    worksheet.Cells[startRow + i, startCol + j].PutValue(rowArray[j]?.GetValue<string>() ?? "");
                }
            }
        }

        workbook.Save(path);

        return await Task.FromResult($"Data written to range starting at {startCell}: {path}");
    }
}

