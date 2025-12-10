using System.Text.Json.Nodes;
using Aspose.Cells;

namespace AsposeMcpServer.Tools;

public class ExcelClearRangeTool : IAsposeTool
{
    public string Description => "Clear range content, format, or both in Excel";

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
                description = "Cell range (e.g., 'A1:C5')"
            },
            clearContent = new
            {
                type = "boolean",
                description = "Clear cell content (optional, default: true)"
            },
            clearFormat = new
            {
                type = "boolean",
                description = "Clear cell format (optional, default: false)"
            }
        },
        required = new[] { "path", "range" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var sheetIndex = arguments?["sheetIndex"]?.GetValue<int>() ?? 0;
        var range = arguments?["range"]?.GetValue<string>() ?? throw new ArgumentException("range is required");
        var clearContent = arguments?["clearContent"]?.GetValue<bool?>() ?? true;
        var clearFormat = arguments?["clearFormat"]?.GetValue<bool?>() ?? false;

        using var workbook = new Workbook(path);
        if (sheetIndex < 0 || sheetIndex >= workbook.Worksheets.Count)
        {
            throw new ArgumentException($"sheetIndex must be between 0 and {workbook.Worksheets.Count - 1}");
        }

        var worksheet = workbook.Worksheets[sheetIndex];
        var cells = worksheet.Cells;
        var cellRange = cells.CreateRange(range);

        if (clearContent && clearFormat)
        {
            // Clear content and format by iterating through cells
            for (int i = cellRange.FirstRow; i <= cellRange.FirstRow + cellRange.RowCount - 1; i++)
            {
                for (int j = cellRange.FirstColumn; j <= cellRange.FirstColumn + cellRange.ColumnCount - 1; j++)
                {
                    cells[i, j].PutValue("");
                    var defaultStyle = workbook.CreateStyle();
                    cells[i, j].SetStyle(defaultStyle);
                }
            }
        }
        else if (clearContent)
        {
            // Clear content only
            for (int i = cellRange.FirstRow; i <= cellRange.FirstRow + cellRange.RowCount - 1; i++)
            {
                for (int j = cellRange.FirstColumn; j <= cellRange.FirstColumn + cellRange.ColumnCount - 1; j++)
                {
                    cells[i, j].PutValue("");
                }
            }
        }
        else if (clearFormat)
        {
            // Clear format only
            var defaultStyle = workbook.CreateStyle();
            cellRange.ApplyStyle(defaultStyle, new StyleFlag { All = true });
        }

        workbook.Save(path);
        return await Task.FromResult($"Range {range} cleared in sheet {sheetIndex}: {path}");
    }
}

