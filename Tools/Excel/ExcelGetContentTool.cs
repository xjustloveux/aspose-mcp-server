using System.Text.Json.Nodes;
using System.Text.Json;
using Aspose.Cells;

namespace AsposeMcpServer.Tools;

public class ExcelGetContentTool : IAsposeTool
{
    public string Description => "Read data from an Excel file";

    public object InputSchema => new
    {
        type = "object",
        properties = new
        {
            path = new
            {
                type = "string",
                description = "Input file path"
            },
            sheetIndex = new
            {
                type = "number",
                description = "Sheet index to read (0-based, optional)"
            },
            range = new
            {
                type = "string",
                description = "Cell range to read (e.g., 'A1:C10', optional)"
            }
        },
        required = new[] { "path" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var sheetIndex = arguments?["sheetIndex"]?.GetValue<int>() ?? 0;
        var range = arguments?["range"]?.GetValue<string>();

        using var workbook = new Workbook(path);
        var worksheet = workbook.Worksheets[sheetIndex];

        if (!string.IsNullOrEmpty(range))
        {
            var cells = worksheet.Cells;
            var cellRange = cells.CreateRange(range);
            var options = new ExportTableOptions
            {
                ExportColumnName = false
            };
            var data = cellRange.ExportDataTable(options);
            
            var result = new List<List<object>>();
            for (int i = 0; i < data.Rows.Count; i++)
            {
                var row = new List<object>();
                for (int j = 0; j < data.Columns.Count; j++)
                {
                    row.Add(data.Rows[i][j] ?? "");
                }
                result.Add(row);
            }
            
            return await Task.FromResult(JsonSerializer.Serialize(result));
        }
        else
        {
            var maxRow = worksheet.Cells.MaxDataRow + 1;
            var maxCol = worksheet.Cells.MaxDataColumn + 1;
            
            var result = new List<List<object>>();
            for (int i = 0; i < maxRow; i++)
            {
                var row = new List<object>();
                for (int j = 0; j < maxCol; j++)
                {
                    row.Add(worksheet.Cells[i, j].Value ?? "");
                }
                result.Add(row);
            }
            
            return await Task.FromResult(JsonSerializer.Serialize(result));
        }
    }
}

