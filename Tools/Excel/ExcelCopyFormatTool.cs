using System.Text.Json.Nodes;
using Aspose.Cells;

namespace AsposeMcpServer.Tools;

public class ExcelCopyFormatTool : IAsposeTool
{
    public string Description => "Copy cell format from source range to destination range (format painter)";

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
            sourceRange = new
            {
                type = "string",
                description = "Source cell range to copy format from (e.g., 'A1')"
            },
            destRange = new
            {
                type = "string",
                description = "Destination cell range to apply format to (e.g., 'B1:C10')"
            },
            copyValue = new
            {
                type = "boolean",
                description = "Copy cell values as well (default: false, only copy format)"
            }
        },
        required = new[] { "path", "sourceRange", "destRange" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var outputPath = arguments?["outputPath"]?.GetValue<string>() ?? path;
        var sheetIndex = arguments?["sheetIndex"]?.GetValue<int>() ?? 0;
        var sourceRange = arguments?["sourceRange"]?.GetValue<string>() ?? throw new ArgumentException("sourceRange is required");
        var destRange = arguments?["destRange"]?.GetValue<string>() ?? throw new ArgumentException("destRange is required");
        var copyValue = arguments?["copyValue"]?.GetValue<bool>() ?? false;

        using var workbook = new Workbook(path);
        
        if (sheetIndex < 0 || sheetIndex >= workbook.Worksheets.Count)
        {
            throw new ArgumentException($"工作表索引 {sheetIndex} 超出範圍 (共有 {workbook.Worksheets.Count} 個工作表)");
        }

        var worksheet = workbook.Worksheets[sheetIndex];
        var cells = worksheet.Cells;
        
        var sourceCellRange = cells.CreateRange(sourceRange);
        var destCellRange = cells.CreateRange(destRange);

        // Copy format using PasteOptions
        var pasteOptions = new PasteOptions();
        pasteOptions.PasteType = copyValue ? PasteType.All : PasteType.Formats;
        pasteOptions.SkipBlanks = false;

        destCellRange.Copy(sourceCellRange, pasteOptions);

        workbook.Save(outputPath);

        var result = $"成功複製格式";
        if (copyValue)
        {
            result += "和值";
        }
        result += $"\n來源範圍: {sourceRange}\n目標範圍: {destRange}\n輸出: {outputPath}";

        return await Task.FromResult(result);
    }
}

