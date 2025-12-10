using System.Text.Json.Nodes;
using Aspose.Cells;

namespace AsposeMcpServer.Tools;

public class ExcelSetPrintTitlesTool : IAsposeTool
{
    public string Description => "Set print titles (rows/columns to repeat on each page) in Excel";

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
            rows = new
            {
                type = "string",
                description = "Rows to repeat (e.g., '1:1' for first row, optional)"
            },
            columns = new
            {
                type = "string",
                description = "Columns to repeat (e.g., 'A:A' for first column, optional)"
            },
            clearTitles = new
            {
                type = "boolean",
                description = "Clear print titles (optional, default: false)"
            }
        },
        required = new[] { "path" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var sheetIndex = arguments?["sheetIndex"]?.GetValue<int?>() ?? 0;
        var rows = arguments?["rows"]?.GetValue<string>();
        var columns = arguments?["columns"]?.GetValue<string>();
        var clearTitles = arguments?["clearTitles"]?.GetValue<bool?>() ?? false;

        using var workbook = new Workbook(path);
        if (sheetIndex < 0 || sheetIndex >= workbook.Worksheets.Count)
        {
            throw new ArgumentException($"sheetIndex must be between 0 and {workbook.Worksheets.Count - 1}");
        }

        var worksheet = workbook.Worksheets[sheetIndex];

        if (clearTitles)
        {
            worksheet.PageSetup.PrintTitleRows = "";
            worksheet.PageSetup.PrintTitleColumns = "";
        }
        else
        {
            if (!string.IsNullOrEmpty(rows))
            {
                worksheet.PageSetup.PrintTitleRows = rows;
            }
            if (!string.IsNullOrEmpty(columns))
            {
                worksheet.PageSetup.PrintTitleColumns = columns;
            }
        }

        workbook.Save(path);
        return await Task.FromResult($"Print titles updated for sheet {sheetIndex}: {path}");
    }
}

