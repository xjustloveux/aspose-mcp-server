using System.Text.Json.Nodes;
using System.Text;
using Aspose.Cells;

namespace AsposeMcpServer.Tools;

public class ExcelGetFormulaTool : IAsposeTool
{
    public string Description => "Get formulas from cells in an Excel worksheet";

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
                description = "Cell range to get formulas from (e.g., 'A1:C10'). If not provided, gets all formulas in the sheet."
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
        
        if (sheetIndex < 0 || sheetIndex >= workbook.Worksheets.Count)
        {
            throw new ArgumentException($"工作表索引 {sheetIndex} 超出範圍 (共有 {workbook.Worksheets.Count} 個工作表)");
        }

        var worksheet = workbook.Worksheets[sheetIndex];
        var cells = worksheet.Cells;
        var result = new StringBuilder();

        result.AppendLine($"=== 工作表 '{worksheet.Name}' 的公式資訊 ===\n");

        int startRow, endRow, startCol, endCol;

        if (!string.IsNullOrEmpty(range))
        {
            try
            {
                var cellRange = cells.CreateRange(range);
                startRow = cellRange.FirstRow;
                endRow = cellRange.FirstRow + cellRange.RowCount - 1;
                startCol = cellRange.FirstColumn;
                endCol = cellRange.FirstColumn + cellRange.ColumnCount - 1;
            }
            catch (Exception ex)
            {
                throw new ArgumentException($"無效的範圍格式: {range}", ex);
            }
        }
        else
        {
            startRow = 0;
            endRow = worksheet.Cells.MaxDataRow;
            startCol = 0;
            endCol = worksheet.Cells.MaxDataColumn;
        }

        int formulaCount = 0;
        for (int row = startRow; row <= endRow && row <= 10000; row++) // Limit for performance
        {
            for (int col = startCol; col <= endCol && col <= 1000; col++)
            {
                var cell = cells[row, col];
                if (!string.IsNullOrEmpty(cell.Formula))
                {
                    formulaCount++;
                    result.AppendLine($"【{CellsHelper.CellIndexToName(row, col)}】");
                    result.AppendLine($"公式: {cell.Formula}");
                    result.AppendLine($"值: {cell.Value ?? "(計算中)"}");
                    result.AppendLine();
                }
            }
        }

        if (formulaCount == 0)
        {
            result.AppendLine("未找到公式");
        }
        else
        {
            result.Insert(0, $"總公式數: {formulaCount}\n\n");
        }

        return await Task.FromResult(result.ToString());
    }
}

