using System.Text.Json.Nodes;
using System.Text;
using Aspose.Cells;

namespace AsposeMcpServer.Tools;

public class ExcelGetUsedRangeTool : IAsposeTool
{
    public string Description => "Get used range (data range) in Excel worksheet";

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
            }
        },
        required = new[] { "path" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var sheetIndex = arguments?["sheetIndex"]?.GetValue<int?>() ?? 0;

        using var workbook = new Workbook(path);
        if (sheetIndex < 0 || sheetIndex >= workbook.Worksheets.Count)
        {
            throw new ArgumentException($"sheetIndex must be between 0 and {workbook.Worksheets.Count - 1}");
        }

        var worksheet = workbook.Worksheets[sheetIndex];
        var cells = worksheet.Cells;

        var sb = new StringBuilder();
        sb.AppendLine($"Used Range for Sheet '{worksheet.Name}':");
        sb.AppendLine($"  First Row: {cells.MinDataRow}");
        sb.AppendLine($"  Last Row: {cells.MaxDataRow}");
        sb.AppendLine($"  First Column: {cells.MinDataColumn}");
        sb.AppendLine($"  Last Column: {cells.MaxDataColumn}");
        
        if (cells.MaxDataRow >= 0 && cells.MaxDataColumn >= 0)
        {
            var firstCell = CellsHelper.CellIndexToName(cells.MinDataRow, cells.MinDataColumn);
            var lastCell = CellsHelper.CellIndexToName(cells.MaxDataRow, cells.MaxDataColumn);
            sb.AppendLine($"  Range: {firstCell}:{lastCell}");
            sb.AppendLine($"  Total Rows: {cells.MaxDataRow - cells.MinDataRow + 1}");
            sb.AppendLine($"  Total Columns: {cells.MaxDataColumn - cells.MinDataColumn + 1}");
        }
        else
        {
            sb.AppendLine("  No data found");
        }

        return await Task.FromResult(sb.ToString());
    }
}

