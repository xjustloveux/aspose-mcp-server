using System.Text.Json.Nodes;
using Aspose.Cells;

namespace AsposeMcpServer.Tools;

public class ExcelAddFormulaTool : IAsposeTool
{
    public string Description => "Add a formula to an Excel cell";

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
                description = "Target cell (e.g., 'A1')"
            },
            formula = new
            {
                type = "string",
                description = "Formula (e.g., '=SUM(A1:A10)')"
            }
        },
        required = new[] { "path", "cell", "formula" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var sheetIndex = arguments?["sheetIndex"]?.GetValue<int>() ?? 0;
        var cell = arguments?["cell"]?.GetValue<string>() ?? throw new ArgumentException("cell is required");
        var formula = arguments?["formula"]?.GetValue<string>() ?? throw new ArgumentException("formula is required");

        using var workbook = new Workbook(path);
        var worksheet = workbook.Worksheets[sheetIndex];
        
        worksheet.Cells[cell].Formula = formula;
        workbook.CalculateFormula();
        
        workbook.Save(path);

        return await Task.FromResult($"Formula added to cell {cell}: {formula}");
    }
}

