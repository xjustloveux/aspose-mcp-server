using System.Text.Json.Nodes;
using Aspose.Cells;

namespace AsposeMcpServer.Tools;

public class ExcelCalculateFormulaTool : IAsposeTool
{
    public string Description => "Calculate formulas in Excel workbook";

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
            outputPath = new
            {
                type = "string",
                description = "Output file path (if not provided, overwrites input)"
            },
            sheetIndex = new
            {
                type = "number",
                description = "Sheet index (0-based, optional, if not provided calculates all sheets)"
            },
            cell = new
            {
                type = "string",
                description = "Specific cell to calculate (e.g., 'A1', optional, if not provided calculates all formulas)"
            }
        },
        required = new[] { "path" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var outputPath = arguments?["outputPath"]?.GetValue<string>() ?? path;
        var sheetIndex = arguments?["sheetIndex"]?.GetValue<int?>();
        var cell = arguments?["cell"]?.GetValue<string>();

        using var workbook = new Workbook(path);
        
        if (!string.IsNullOrEmpty(cell) && sheetIndex.HasValue)
        {
            // Calculate specific cell
            if (sheetIndex.Value < 0 || sheetIndex.Value >= workbook.Worksheets.Count)
            {
                throw new ArgumentException($"工作表索引 {sheetIndex.Value} 超出範圍 (共有 {workbook.Worksheets.Count} 個工作表)");
            }
            
            var worksheet = workbook.Worksheets[sheetIndex.Value];
            var cellObj = worksheet.Cells[cell];
            
            if (!string.IsNullOrEmpty(cellObj.Formula))
            {
                // Recalculate the cell by setting its value
                var oldValue = cellObj.Value;
                cellObj.PutValue(cellObj.Formula);
            }
        }
        else if (sheetIndex.HasValue)
        {
            // Calculate all formulas in specific sheet
            if (sheetIndex.Value < 0 || sheetIndex.Value >= workbook.Worksheets.Count)
            {
                throw new ArgumentException($"工作表索引 {sheetIndex.Value} 超出範圍 (共有 {workbook.Worksheets.Count} 個工作表)");
            }
            
            // Calculate formulas in the specific sheet by calculating the entire workbook
            // Aspose.Cells doesn't have sheet-level CalculateFormula, so we calculate all
            workbook.CalculateFormula();
        }
        else
        {
            // Calculate all formulas in all sheets
            workbook.CalculateFormula();
        }
        
        workbook.Save(outputPath);
        
        var result = "公式計算完成\n";
        if (sheetIndex.HasValue)
        {
            result += $"工作表: {workbook.Worksheets[sheetIndex.Value].Name}\n";
        }
        else
        {
            result += "範圍: 所有工作表\n";
        }
        if (!string.IsNullOrEmpty(cell))
        {
            result += $"單元格: {cell}\n";
        }
        result += $"輸出: {outputPath}";
        
        return await Task.FromResult(result);
    }
}
