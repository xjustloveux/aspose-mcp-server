using System.Text.Json.Nodes;
using Aspose.Cells;

namespace AsposeMcpServer.Tools;

public class ExcelWriteCellTool : IAsposeTool
{
    public string Description => "Write data to a specific cell in Excel";

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
            value = new
            {
                type = "string",
                description = "Value to write"
            }
        },
        required = new[] { "path", "cell", "value" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var sheetIndex = arguments?["sheetIndex"]?.GetValue<int>() ?? 0;
        var cell = arguments?["cell"]?.GetValue<string>() ?? throw new ArgumentException("cell is required");
        var value = arguments?["value"]?.GetValue<string>() ?? throw new ArgumentException("value is required");

        using var workbook = new Workbook(path);
        if (sheetIndex < 0 || sheetIndex >= workbook.Worksheets.Count)
        {
            throw new ArgumentException($"sheetIndex must be between 0 and {workbook.Worksheets.Count - 1}");
        }

        var worksheet = workbook.Worksheets[sheetIndex];
        worksheet.Cells[cell].PutValue(value);
        
        workbook.Save(path);

        return await Task.FromResult($"Cell {cell} updated in sheet {sheetIndex}: {path}");
    }
}

