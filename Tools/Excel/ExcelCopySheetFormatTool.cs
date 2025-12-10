using System.Text.Json.Nodes;
using Aspose.Cells;

namespace AsposeMcpServer.Tools;

public class ExcelCopySheetFormatTool : IAsposeTool
{
    public string Description => "Copy format from one sheet to another in Excel";

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
            sourceSheetIndex = new
            {
                type = "number",
                description = "Source sheet index (0-based)"
            },
            targetSheetIndex = new
            {
                type = "number",
                description = "Target sheet index (0-based)"
            },
            copyColumnWidths = new
            {
                type = "boolean",
                description = "Copy column widths (optional, default: true)"
            },
            copyRowHeights = new
            {
                type = "boolean",
                description = "Copy row heights (optional, default: true)"
            }
        },
        required = new[] { "path", "sourceSheetIndex", "targetSheetIndex" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var sourceSheetIndex = arguments?["sourceSheetIndex"]?.GetValue<int>() ?? throw new ArgumentException("sourceSheetIndex is required");
        var targetSheetIndex = arguments?["targetSheetIndex"]?.GetValue<int>() ?? throw new ArgumentException("targetSheetIndex is required");
        var copyColumnWidths = arguments?["copyColumnWidths"]?.GetValue<bool?>() ?? true;
        var copyRowHeights = arguments?["copyRowHeights"]?.GetValue<bool?>() ?? true;

        using var workbook = new Workbook(path);
        if (sourceSheetIndex < 0 || sourceSheetIndex >= workbook.Worksheets.Count)
        {
            throw new ArgumentException($"sourceSheetIndex must be between 0 and {workbook.Worksheets.Count - 1}");
        }
        if (targetSheetIndex < 0 || targetSheetIndex >= workbook.Worksheets.Count)
        {
            throw new ArgumentException($"targetSheetIndex must be between 0 and {workbook.Worksheets.Count - 1}");
        }

        var sourceSheet = workbook.Worksheets[sourceSheetIndex];
        var targetSheet = workbook.Worksheets[targetSheetIndex];

        // Copy column widths
        if (copyColumnWidths)
        {
            for (int i = 0; i <= sourceSheet.Cells.MaxDataColumn; i++)
            {
                targetSheet.Cells.SetColumnWidth(i, sourceSheet.Cells.GetColumnWidth(i));
            }
        }

        // Copy row heights
        if (copyRowHeights)
        {
            for (int i = 0; i <= sourceSheet.Cells.MaxDataRow; i++)
            {
                targetSheet.Cells.SetRowHeight(i, sourceSheet.Cells.GetRowHeight(i));
            }
        }

        workbook.Save(path);
        return await Task.FromResult($"Format copied from sheet {sourceSheetIndex} to sheet {targetSheetIndex}: {path}");
    }
}

