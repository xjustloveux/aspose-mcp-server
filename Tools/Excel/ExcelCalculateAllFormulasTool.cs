using System.Text.Json.Nodes;
using Aspose.Cells;

namespace AsposeMcpServer.Tools;

public class ExcelCalculateAllFormulasTool : IAsposeTool
{
    public string Description => "Calculate all formulas in Excel workbook";

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
                description = "Sheet index (0-based, optional, if not provided calculates all sheets)"
            }
        },
        required = new[] { "path" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var sheetIndex = arguments?["sheetIndex"]?.GetValue<int?>();

        using var workbook = new Workbook(path);

        // Calculate formulas for specific sheet or all sheets
        if (sheetIndex.HasValue)
        {
            var worksheet = workbook.Worksheets[sheetIndex.Value];
            if (worksheet != null)
            {
                // Force calculation for the worksheet - CalculateFormula doesn't take parameters
                // Instead, we calculate all formulas in workbook which will include this sheet
                workbook.CalculateFormula();
            }
        }
        else
        {
            // Calculate all formulas in workbook
            workbook.CalculateFormula();
        }
        
        // Save the workbook to persist calculated values
        workbook.Save(path);
        return await Task.FromResult(sheetIndex.HasValue
            ? $"Formulas calculated for sheet {sheetIndex.Value}: {path}"
            : $"All formulas calculated: {path}");
    }
}

