using System.Text.Json.Nodes;
using System.Text;
using Aspose.Cells;

namespace AsposeMcpServer.Tools;

public class ExcelGetArrayFormulaTool : IAsposeTool
{
    public string Description => "Get array formula from Excel range";

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
                description = "Cell reference (e.g., 'A1') - any cell in the array formula range"
            }
        },
        required = new[] { "path", "cell" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var sheetIndex = arguments?["sheetIndex"]?.GetValue<int?>() ?? 0;
        var cell = arguments?["cell"]?.GetValue<string>() ?? throw new ArgumentException("cell is required");

        using var workbook = new Workbook(path);
        if (sheetIndex < 0 || sheetIndex >= workbook.Worksheets.Count)
        {
            throw new ArgumentException($"sheetIndex must be between 0 and {workbook.Worksheets.Count - 1}");
        }

        var worksheet = workbook.Worksheets[sheetIndex];
        var cellObj = worksheet.Cells[cell];

        var sb = new StringBuilder();
        sb.AppendLine($"Cell: {cell}");

        // Check if cell is part of an array formula
        var formula = cellObj.Formula;
        if (!string.IsNullOrEmpty(formula) && formula.StartsWith("{"))
        {
            sb.AppendLine($"Array Formula: {formula}");
            // Try to find the array range
            var range = worksheet.Cells.CreateRange(cell);
            sb.AppendLine($"Array Range: {range.Name}");
        }
        else
        {
            sb.AppendLine("No array formula found in this cell");
        }

        return await Task.FromResult(sb.ToString());
    }
}

