using System.Text.Json.Nodes;
using System.Text;
using Aspose.Cells;

namespace AsposeMcpServer.Tools;

public class ExcelGetCellFormatTool : IAsposeTool
{
    public string Description => "Get detailed cell format information from an Excel worksheet";

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
                description = "Cell address (e.g., 'A1') or range (e.g., 'A1:C3')"
            }
        },
        required = new[] { "path", "cell" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var sheetIndex = arguments?["sheetIndex"]?.GetValue<int>() ?? 0;
        var cell = arguments?["cell"]?.GetValue<string>() ?? throw new ArgumentException("cell is required");

        using var workbook = new Workbook(path);
        
        if (sheetIndex < 0 || sheetIndex >= workbook.Worksheets.Count)
        {
            throw new ArgumentException($"工作表索引 {sheetIndex} 超出範圍 (共有 {workbook.Worksheets.Count} 個工作表)");
        }

        var worksheet = workbook.Worksheets[sheetIndex];
        var cells = worksheet.Cells;
        var result = new StringBuilder();

        result.AppendLine($"=== 工作表 '{worksheet.Name}' 的單元格格式資訊 ===\n");

        try
        {
            var cellRange = cells.CreateRange(cell);
            var startRow = cellRange.FirstRow;
            var endRow = cellRange.FirstRow + cellRange.RowCount - 1;
            var startCol = cellRange.FirstColumn;
            var endCol = cellRange.FirstColumn + cellRange.ColumnCount - 1;

            for (int row = startRow; row <= endRow; row++)
            {
                for (int col = startCol; col <= endCol; col++)
                {
                    var cellObj = cells[row, col];
                    var style = cellObj.GetStyle();
                    
                    result.AppendLine($"【單元格 {CellsHelper.CellIndexToName(row, col)}】");
                    result.AppendLine($"值: {cellObj.Value ?? "(空)"}");
                    result.AppendLine($"公式: {cellObj.Formula ?? "(無)"}");
                    result.AppendLine($"數據類型: {cellObj.Type}");
                    result.AppendLine();
                    
                    result.AppendLine("格式資訊:");
                    result.AppendLine($"  字型: {style.Font.Name}, 大小: {style.Font.Size}");
                    result.AppendLine($"  粗體: {style.Font.IsBold}, 斜體: {style.Font.IsItalic}");
                    result.AppendLine($"  底線: {style.Font.Underline}, 刪除線: {style.Font.IsStrikeout}");
                    result.AppendLine($"  字型顏色: {style.Font.Color}");
                    result.AppendLine($"  背景色: {style.BackgroundColor}");
                    result.AppendLine($"  水平對齊: {style.HorizontalAlignment}");
                    result.AppendLine($"  垂直對齊: {style.VerticalAlignment}");
                    result.AppendLine($"  文字換行: {style.IsTextWrapped}");
                    result.AppendLine($"  數字格式: {style.Number}");
                    result.AppendLine($"  邊框: 上={style.Borders[BorderType.TopBorder].LineStyle}, 下={style.Borders[BorderType.BottomBorder].LineStyle}, 左={style.Borders[BorderType.LeftBorder].LineStyle}, 右={style.Borders[BorderType.RightBorder].LineStyle}");
                    result.AppendLine();
                }
            }
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException($"無法讀取單元格格式: {ex.Message}", ex);
        }

        return await Task.FromResult(result.ToString());
    }
}

