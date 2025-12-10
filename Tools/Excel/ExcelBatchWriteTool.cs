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

        foreach (var item in dataArray)
        {
            if (item is JsonObject dataItem)
            {
                var cellRef = dataItem["cell"]?.GetValue<string>();
                var value = dataItem["value"]?.GetValue<string>();

                if (!string.IsNullOrEmpty(cellRef))
                {
                    cells[cellRef].PutValue(value ?? "");
                    writtenCount++;
                }
            }
        }

        workbook.Save(outputPath);

        return await Task.FromResult($"成功批量寫入 {writtenCount} 個單元格\n工作表: {worksheet.Name}\n輸出: {outputPath}");
    }
}

