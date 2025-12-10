using System.Text.Json.Nodes;
using Aspose.Cells;

namespace AsposeMcpServer.Tools;

public class ExcelSetPrintAreaTool : IAsposeTool
{
    public string Description => "Set print area in Excel worksheet";

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
                description = "Print area range (e.g., 'A1:D10', optional, if not provided clears print area)"
            },
            clearPrintArea = new
            {
                type = "boolean",
                description = "Clear print area (optional, default: false)"
            }
        },
        required = new[] { "path" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var sheetIndex = arguments?["sheetIndex"]?.GetValue<int?>() ?? 0;
        var range = arguments?["range"]?.GetValue<string>();
        var clearPrintArea = arguments?["clearPrintArea"]?.GetValue<bool?>() ?? false;

        using var workbook = new Workbook(path);
        if (sheetIndex < 0 || sheetIndex >= workbook.Worksheets.Count)
        {
            throw new ArgumentException($"sheetIndex must be between 0 and {workbook.Worksheets.Count - 1}");
        }

        var worksheet = workbook.Worksheets[sheetIndex];

        if (clearPrintArea)
        {
            worksheet.PageSetup.PrintArea = "";
        }
        else if (!string.IsNullOrEmpty(range))
        {
            worksheet.PageSetup.PrintArea = range;
        }
        else
        {
            throw new ArgumentException("Either range or clearPrintArea must be provided");
        }

        workbook.Save(path);
        return await Task.FromResult(clearPrintArea
            ? $"Print area cleared for sheet {sheetIndex}: {path}"
            : $"Print area set to {range} for sheet {sheetIndex}: {path}");
    }
}

