using System.Text.Json.Nodes;
using Aspose.Cells;

namespace AsposeMcpServer.Tools;

public class ExcelSetArrayFormulaTool : IAsposeTool
{
    public string Description => "Set array formula in Excel range";

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
                description = "Range for array formula (e.g., 'A1:C3')"
            },
            formula = new
            {
                type = "string",
                description = "Array formula (e.g., '{=SUM(A1:A3*B1:B3)}')"
            }
        },
        required = new[] { "path", "range", "formula" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var sheetIndex = arguments?["sheetIndex"]?.GetValue<int?>() ?? 0;
        var range = arguments?["range"]?.GetValue<string>() ?? throw new ArgumentException("range is required");
        var formula = arguments?["formula"]?.GetValue<string>() ?? throw new ArgumentException("formula is required");

        using var workbook = new Workbook(path);
        if (sheetIndex < 0 || sheetIndex >= workbook.Worksheets.Count)
        {
            throw new ArgumentException($"sheetIndex must be between 0 and {workbook.Worksheets.Count - 1}");
        }

        var worksheet = workbook.Worksheets[sheetIndex];
        var rangeObj = worksheet.Cells.CreateRange(range);

        // Remove curly braces if present (they're added automatically)
        var cleanFormula = formula.TrimStart('{').TrimEnd('}');
        // Set array formula by setting formula in first cell and copying to range
        var firstCell = worksheet.Cells[rangeObj.FirstRow, rangeObj.FirstColumn];
        firstCell.Formula = cleanFormula;
        // Copy formula to all cells in range
        for (int i = rangeObj.FirstRow; i <= rangeObj.FirstRow + rangeObj.RowCount - 1; i++)
        {
            for (int j = rangeObj.FirstColumn; j <= rangeObj.FirstColumn + rangeObj.ColumnCount - 1; j++)
            {
                if (i != rangeObj.FirstRow || j != rangeObj.FirstColumn)
                {
                    worksheet.Cells[i, j].Formula = cleanFormula;
                }
            }
        }

        workbook.Save(path);
        return await Task.FromResult($"Array formula set in range {range}: {path}");
    }
}

