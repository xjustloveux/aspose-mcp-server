using System.Text.Json.Nodes;
using Aspose.Cells;

namespace AsposeMcpServer.Tools;

public class ExcelBatchWriteTool : IAsposeTool
{
    public string Description => "Batch write data to multiple cells in an Excel worksheet";

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
                description = "Sheet index (0-based, optional, default: 0)"
            },
            data = new
            {
                type = "array",
                description = "Array of objects with 'cell' (e.g., 'A1') and 'value' properties",
                items = new
                {
                    type = "object",
                    properties = new
                    {
                        cell = new { type = "string" },
                        value = new { type = "string" }
                    },
                    required = new[] { "cell", "value" }
                }
            }
        },
        required = new[] { "path", "data" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var outputPath = arguments?["outputPath"]?.GetValue<string>() ?? path;
        var sheetIndex = arguments?["sheetIndex"]?.GetValue<int>() ?? 0;
        var dataArray = arguments?["data"]?.AsArray() ?? throw new ArgumentException("data is required");

        using var workbook = new Workbook(path);
        
        if (sheetIndex < 0 || sheetIndex >= workbook.Worksheets.Count)
        {
            throw new ArgumentException($"工作表索引 {sheetIndex} 超出範圍 (共有 {workbook.Worksheets.Count} 個工作表)");
        }

        var worksheet = workbook.Worksheets[sheetIndex];
        var cells = worksheet.Cells;
        int writtenCount = 0;

        // Separate data and formulas to ensure data is written before formulas
        var dataItems = new List<(string cellRef, string value)>();
        var formulaItems = new List<(string cellRef, string formula)>();

        foreach (var item in dataArray)
        {
            if (item is JsonObject dataItem)
            {
                var cellRef = dataItem["cell"]?.GetValue<string>();
                var value = dataItem["value"]?.GetValue<string>();

                if (!string.IsNullOrEmpty(cellRef))
                {
                    // Auto-detect formulas: if value starts with "=", treat it as a formula
                    if (!string.IsNullOrEmpty(value) && value.StartsWith("="))
                    {
                        formulaItems.Add((cellRef, value));
                    }
                    else
                    {
                        dataItems.Add((cellRef, value ?? ""));
                    }
                }
            }
        }

        // First, write all data values (non-formulas)
        // Try to parse numeric values to ensure they're stored as numbers, not strings
        foreach (var (cellRef, value) in dataItems)
        {
            var cell = cells[cellRef];
            
            // Try to parse as number if possible
            if (double.TryParse(value, out double numValue))
            {
                cell.PutValue(numValue);
            }
            else if (DateTime.TryParse(value, out DateTime dateValue))
            {
                cell.PutValue(dateValue);
            }
            else
            {
                cell.PutValue(value);
            }
            
            writtenCount++;
        }

        // Then, write all formulas (after data is written)
        foreach (var (cellRef, formula) in formulaItems)
        {
            var cell = cells[cellRef];
            // Set formula explicitly
            cell.Formula = formula;
            writtenCount++;
        }

        // Force recalculation: calculate formulas multiple times to ensure dependencies are resolved
        workbook.CalculateFormula();
        
        // Also try to force calculation by accessing cell values
        foreach (var (cellRef, formula) in formulaItems)
        {
            var cell = cells[cellRef];
            // Force evaluation by accessing the value
            _ = cell.Value;
        }
        
        // Calculate again after accessing values
        workbook.CalculateFormula();
        
        workbook.Save(outputPath);

        return await Task.FromResult($"成功批量寫入 {writtenCount} 個單元格\n工作表: {worksheet.Name}\n輸出: {outputPath}");
    }
}

