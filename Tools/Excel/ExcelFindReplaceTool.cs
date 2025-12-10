using System.Text.Json.Nodes;
using Aspose.Cells;

namespace AsposeMcpServer.Tools;

public class ExcelFindReplaceTool : IAsposeTool
{
    public string Description => "Find and replace text in Excel cells";

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
            findText = new
            {
                type = "string",
                description = "Text to find"
            },
            replaceText = new
            {
                type = "string",
                description = "Text to replace with"
            },
            sheetIndex = new
            {
                type = "number",
                description = "Sheet index (0-based, optional, if not provided searches all sheets)"
            },
            matchCase = new
            {
                type = "boolean",
                description = "Match case (default: false)"
            },
            matchEntireCell = new
            {
                type = "boolean",
                description = "Match entire cell content (default: false)"
            }
        },
        required = new[] { "path", "findText", "replaceText" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var outputPath = arguments?["outputPath"]?.GetValue<string>() ?? path;
        var findText = arguments?["findText"]?.GetValue<string>() ?? throw new ArgumentException("findText is required");
        var replaceText = arguments?["replaceText"]?.GetValue<string>() ?? throw new ArgumentException("replaceText is required");
        var sheetIndex = arguments?["sheetIndex"]?.GetValue<int?>();
        var matchCase = arguments?["matchCase"]?.GetValue<bool>() ?? false;
        var matchEntireCell = arguments?["matchEntireCell"]?.GetValue<bool>() ?? false;

        using var workbook = new Workbook(path);
        var totalReplacements = 0;
        var lookAt = matchEntireCell ? LookAtType.EntireContent : LookAtType.Contains;
        
        if (sheetIndex.HasValue)
        {
            if (sheetIndex.Value < 0 || sheetIndex.Value >= workbook.Worksheets.Count)
            {
                throw new ArgumentException($"工作表索引 {sheetIndex.Value} 超出範圍 (共有 {workbook.Worksheets.Count} 個工作表)");
            }
            
            var worksheet = workbook.Worksheets[sheetIndex.Value];
            totalReplacements += ReplaceInWorksheet(worksheet, findText, replaceText, matchCase, lookAt);
        }
        else
        {
            // Search all sheets
            for (int i = 0; i < workbook.Worksheets.Count; i++)
            {
                var worksheet = workbook.Worksheets[i];
                totalReplacements += ReplaceInWorksheet(worksheet, findText, replaceText, matchCase, lookAt);
            }
        }
        
        workbook.Save(outputPath);
        
        return await Task.FromResult($"查找替換完成\n查找: '{findText}'\n替換為: '{replaceText}'\n總替換數: {totalReplacements}\n輸出: {outputPath}");
    }

    private int ReplaceInWorksheet(Worksheet worksheet, string findText, string replaceText, bool matchCase, LookAtType lookAt)
    {
        var options = new FindOptions
        {
            CaseSensitive = matchCase,
            LookAtType = lookAt
        };

        var replacements = 0;
        var cell = worksheet.Cells.Find(findText, null, options);
        while (cell != null)
        {
            cell.PutValue(replaceText);
            replacements++;
            cell = worksheet.Cells.Find(findText, cell, options);
        }

        return replacements;
    }
}

