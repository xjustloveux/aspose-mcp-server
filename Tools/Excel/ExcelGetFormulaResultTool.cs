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

        var worksheet = workbook.Worksheets[sheetIndex];
        var cellObj = worksheet.Cells[cell];

        if (calculateBeforeRead)
        {
            // Calculate formulas before reading
            workbook.CalculateFormula();
        }

        var result = $"Cell: {cell}\n";
        result += $"Formula: {cellObj.Formula ?? "(none)"}\n";
        
        // Get the calculated value - for formula cells, Value should contain the calculated result
        object? calculatedValue = cellObj.Value;
        
        // If cell has formula but value is empty or null, try to get the display value
        if (!string.IsNullOrEmpty(cellObj.Formula))
        {
            if (calculatedValue == null || (calculatedValue is string str && string.IsNullOrEmpty(str)))
            {
                // Formula exists but not calculated - try to get display value
                calculatedValue = cellObj.DisplayStringValue;
                if (string.IsNullOrEmpty(calculatedValue?.ToString()))
                {
                    calculatedValue = cellObj.Formula; // Fallback to formula text
                }
            }
        }
        
        result += $"Calculated Value: {calculatedValue ?? "(empty)"}\n";
        result += $"Value Type: {cellObj.Type}";

        return await Task.FromResult(result);
    }
}

