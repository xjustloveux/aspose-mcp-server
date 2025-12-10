using System.Text.Json.Nodes;
using Aspose.Cells;

namespace AsposeMcpServer.Tools;

public class ExcelEditCellTool : IAsposeTool
{
    public string Description => "Edit cell value and optionally format in Excel";

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
                description = "New value to set (optional)"
            },
            formula = new
            {
                type = "string",
                description = "Formula to set (optional, overrides value)"
            },
            clearValue = new
            {
                type = "boolean",
                description = "Clear cell value (optional)"
            }
        },
        required = new[] { "path", "cell" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var sheetIndex = arguments?["sheetIndex"]?.GetValue<int>() ?? 0;
        var cell = arguments?["cell"]?.GetValue<string>() ?? throw new ArgumentException("cell is required");
        var value = arguments?["value"]?.GetValue<string>();
        var formula = arguments?["formula"]?.GetValue<string>();
        var clearValue = arguments?["clearValue"]?.GetValue<bool?>() ?? false;

        using var workbook = new Workbook(path);
        if (sheetIndex < 0 || sheetIndex >= workbook.Worksheets.Count)
        {
            throw new ArgumentException($"sheetIndex must be between 0 and {workbook.Worksheets.Count - 1}");
        }

        var worksheet = workbook.Worksheets[sheetIndex];
        var cellObj = worksheet.Cells[cell];

        if (clearValue)
        {
            cellObj.PutValue("");
        }
        else if (!string.IsNullOrEmpty(formula))
        {
            cellObj.Formula = formula;
        }
        else if (!string.IsNullOrEmpty(value))
        {
            cellObj.PutValue(value);
        }
        else
        {
            throw new ArgumentException("Either value, formula, or clearValue must be provided");
        }

        workbook.Save(path);
        return await Task.FromResult($"Cell {cell} edited in sheet {sheetIndex}: {path}");
    }
}

