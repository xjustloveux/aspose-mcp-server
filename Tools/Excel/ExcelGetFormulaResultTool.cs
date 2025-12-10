using System.Text.Json.Nodes;
using Aspose.Cells;

namespace AsposeMcpServer.Tools;

public class ExcelGetFormulaResultTool : IAsposeTool
{
    public string Description => "Get calculated formula result from Excel cell";

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
            calculateBeforeRead = new
            {
                type = "boolean",
                description = "Calculate formulas before reading (optional, default: true)"
            }
        },
        required = new[] { "path", "cell" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var sheetIndex = arguments?["sheetIndex"]?.GetValue<int>() ?? 0;
        var cell = arguments?["cell"]?.GetValue<string>() ?? throw new ArgumentException("cell is required");
        var calculateBeforeRead = arguments?["calculateBeforeRead"]?.GetValue<bool?>() ?? true;

        using var workbook = new Workbook(path);
        if (sheetIndex < 0 || sheetIndex >= workbook.Worksheets.Count)
        {
            throw new ArgumentException($"sheetIndex must be between 0 and {workbook.Worksheets.Count - 1}");
        }

        if (calculateBeforeRead)
        {
            workbook.CalculateFormula();
        }

        var worksheet = workbook.Worksheets[sheetIndex];
        var cellObj = worksheet.Cells[cell];

        var result = $"Cell: {cell}\n";
        result += $"Formula: {cellObj.Formula ?? "(none)"}\n";
        result += $"Calculated Value: {cellObj.Value ?? "(empty)"}\n";
        result += $"Value Type: {cellObj.Type}";

        return await Task.FromResult(result);
    }
}

