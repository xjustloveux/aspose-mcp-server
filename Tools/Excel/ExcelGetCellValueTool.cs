using System.Text.Json.Nodes;
using Aspose.Cells;

namespace AsposeMcpServer.Tools;

public class ExcelGetCellValueTool : IAsposeTool
{
    public string Description => "Get cell value, formula, and type from Excel";

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
            includeFormula = new
            {
                type = "boolean",
                description = "Include formula if present (optional, default: true)"
            },
            includeFormat = new
            {
                type = "boolean",
                description = "Include format information (optional, default: false)"
            }
        },
        required = new[] { "path", "cell" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var sheetIndex = arguments?["sheetIndex"]?.GetValue<int>() ?? 0;
        var cell = arguments?["cell"]?.GetValue<string>() ?? throw new ArgumentException("cell is required");
        var includeFormula = arguments?["includeFormula"]?.GetValue<bool?>() ?? true;
        var includeFormat = arguments?["includeFormat"]?.GetValue<bool?>() ?? false;

        using var workbook = new Workbook(path);
        if (sheetIndex < 0 || sheetIndex >= workbook.Worksheets.Count)
        {
            throw new ArgumentException($"sheetIndex must be between 0 and {workbook.Worksheets.Count - 1}");
        }

        var worksheet = workbook.Worksheets[sheetIndex];
        var cellObj = worksheet.Cells[cell];

        var result = new System.Text.StringBuilder();
        result.AppendLine($"Cell: {cell}");
        result.AppendLine($"Value: {cellObj.Value ?? "(empty)"}");
        result.AppendLine($"Value Type: {cellObj.Type}");

        if (includeFormula && !string.IsNullOrEmpty(cellObj.Formula))
        {
            result.AppendLine($"Formula: {cellObj.Formula}");
        }

        if (includeFormat)
        {
            var style = cellObj.GetStyle();
            result.AppendLine($"Format:");
            result.AppendLine($"  Font: {style.Font.Name}, Size: {style.Font.Size}");
            result.AppendLine($"  Bold: {style.Font.IsBold}, Italic: {style.Font.IsItalic}");
            result.AppendLine($"  Background Color: {style.ForegroundColor}");
            result.AppendLine($"  Number Format: {style.Number}");
        }

        return await Task.FromResult(result.ToString());
    }
}

