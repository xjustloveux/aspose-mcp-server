using System.Text.Json.Nodes;
using System.Text;
using Aspose.Cells;

namespace AsposeMcpServer.Tools;

public class ExcelGetMergedCellsTool : IAsposeTool
{
    public string Description => "Get all merged cells information from an Excel worksheet";

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
        var sheetIndex = arguments?["sheetIndex"]?.GetValue<int>() ?? 0;

        using var workbook = new Workbook(path);
        
        if (sheetIndex < 0 || sheetIndex >= workbook.Worksheets.Count)
        {
            throw new ArgumentException($"工作表索引 {sheetIndex} 超出範圍 (共有 {workbook.Worksheets.Count} 個工作表)");
        }

        var worksheet = workbook.Worksheets[sheetIndex];
        var mergedCells = worksheet.Cells.MergedCells;
        if (mergedCells == null)
        {
            throw new InvalidOperationException($"無法取得合併單元格資訊：{worksheet.Name}");
        }
        var result = new StringBuilder();

        result.AppendLine($"=== 工作表 '{worksheet.Name}' 的合併單元格資訊 ===\n");
        result.AppendLine($"總合併區域數: {mergedCells.Count}\n");

        if (mergedCells.Count == 0)
        {
            result.AppendLine("未找到合併單元格");
            return await Task.FromResult(result.ToString());
        }

        for (int i = 0; i < mergedCells.Count; i++)
        {
            var mergedCellName = mergedCells[i]?.ToString();
            if (string.IsNullOrWhiteSpace(mergedCellName)) continue;
            // Parse the merged cell range (format: "A1:B2")
            var parts = mergedCellName.Split(':');
            if (parts.Length == 2)
            {
                var startCell = parts[0];
                var endCell = parts[1];
                
                result.AppendLine($"【合併區域 {i}】");
                result.AppendLine($"範圍: {startCell}:{endCell}");
                
                // Get cell indices
                int startRow, startCol, endRow, endCol;
                CellsHelper.CellNameToIndex(startCell, out startRow, out startCol);
                CellsHelper.CellNameToIndex(endCell, out endRow, out endCol);
                
                result.AppendLine($"行數: {endRow - startRow + 1}");
                result.AppendLine($"列數: {endCol - startCol + 1}");
                
                // Get the value of the merged cell
                var cell = worksheet.Cells[startRow, startCol];
                result.AppendLine($"值: {cell.Value ?? "(空白)"}");
                result.AppendLine();
            }
        }

        return await Task.FromResult(result.ToString());
    }
}

