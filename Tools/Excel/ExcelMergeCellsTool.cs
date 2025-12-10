using System.Text.Json.Nodes;
using Aspose.Cells;

namespace AsposeMcpServer.Tools;

public class ExcelMergeCellsTool : IAsposeTool
{
    public string Description => "Merge or unmerge cells in an Excel worksheet";

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
                description = "Cell range to merge/unmerge (e.g., 'A1:C3')"
            },
            merge = new
            {
                type = "boolean",
                description = "True to merge, false to unmerge (default: true)"
            }
        },
        required = new[] { "path", "range" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var sheetIndex = arguments?["sheetIndex"]?.GetValue<int>() ?? 0;
        var range = arguments?["range"]?.GetValue<string>() ?? throw new ArgumentException("range is required");
        var merge = arguments?["merge"]?.GetValue<bool>() ?? true;

        using var workbook = new Workbook(path);
        var worksheet = workbook.Worksheets[sheetIndex];
        var cells = worksheet.Cells;
        var cellRange = cells.CreateRange(range);

        if (merge)
        {
            cellRange.Merge();
            workbook.Save(path);
            return await Task.FromResult($"範圍 {range} 已合併: {path}");
        }
        else
        {
            cellRange.UnMerge();
            workbook.Save(path);
            return await Task.FromResult($"範圍 {range} 已取消合併: {path}");
        }
    }
}

