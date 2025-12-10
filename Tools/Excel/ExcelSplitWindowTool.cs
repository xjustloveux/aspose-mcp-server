using System.Text.Json.Nodes;
using Aspose.Cells;

namespace AsposeMcpServer.Tools;

public class ExcelSplitWindowTool : IAsposeTool
{
    public string Description => "Split worksheet window in Excel";

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
            splitRow = new
            {
                type = "number",
                description = "Row index to split at (0-based, optional)"
            },
            splitColumn = new
            {
                type = "number",
                description = "Column index to split at (0-based, optional)"
            },
            removeSplit = new
            {
                type = "boolean",
                description = "Remove split (optional, default: false)"
            }
        },
        required = new[] { "path" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var sheetIndex = arguments?["sheetIndex"]?.GetValue<int?>() ?? 0;
        var splitRow = arguments?["splitRow"]?.GetValue<int?>();
        var splitColumn = arguments?["splitColumn"]?.GetValue<int?>();
        var removeSplit = arguments?["removeSplit"]?.GetValue<bool?>() ?? false;

        using var workbook = new Workbook(path);
        if (sheetIndex < 0 || sheetIndex >= workbook.Worksheets.Count)
        {
            throw new ArgumentException($"sheetIndex must be between 0 and {workbook.Worksheets.Count - 1}");
        }

        var worksheet = workbook.Worksheets[sheetIndex];

        if (removeSplit)
        {
            worksheet.RemoveSplit();
        }
        else if (splitRow.HasValue || splitColumn.HasValue)
        {
            // Use FreezePanes to simulate split window
            worksheet.FreezePanes(splitRow ?? 0, splitColumn ?? 0, splitRow ?? 0, splitColumn ?? 0);
        }
        else
        {
            throw new ArgumentException("Either splitRow/splitColumn or removeSplit must be provided");
        }

        workbook.Save(path);
        return await Task.FromResult(removeSplit
            ? $"Window split removed from sheet {sheetIndex}: {path}"
            : $"Window split at row {splitRow ?? 0}, column {splitColumn ?? 0} for sheet {sheetIndex}: {path}");
    }
}

