using System.Text.Json.Nodes;
using Aspose.Cells;

namespace AsposeMcpServer.Tools;

public class ExcelMoveRangeTool : IAsposeTool
{
    public string Description => "Move range from source to destination in Excel";

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
                description = "Source sheet index (0-based, optional, default: 0)"
            },
            sourceRange = new
            {
                type = "string",
                description = "Source range (e.g., 'A1:C5')"
            },
            destSheetIndex = new
            {
                type = "number",
                description = "Destination sheet index (0-based, optional, default: same as source)"
            },
            destCell = new
            {
                type = "string",
                description = "Destination cell (top-left cell, e.g., 'E1')"
            }
        },
        required = new[] { "path", "sourceRange", "destCell" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var sourceSheetIndex = arguments?["sourceSheetIndex"]?.GetValue<int?>() ?? 0;
        var sourceRange = arguments?["sourceRange"]?.GetValue<string>() ?? throw new ArgumentException("sourceRange is required");
        var destSheetIndex = arguments?["destSheetIndex"]?.GetValue<int?>();
        var destCell = arguments?["destCell"]?.GetValue<string>() ?? throw new ArgumentException("destCell is required");

        using var workbook = new Workbook(path);
        if (sourceSheetIndex < 0 || sourceSheetIndex >= workbook.Worksheets.Count)
        {
            throw new ArgumentException($"sourceSheetIndex must be between 0 and {workbook.Worksheets.Count - 1}");
        }

        var destSheetIdx = destSheetIndex ?? sourceSheetIndex;
        if (destSheetIdx < 0 || destSheetIdx >= workbook.Worksheets.Count)
        {
            throw new ArgumentException($"destSheetIndex must be between 0 and {workbook.Worksheets.Count - 1}");
        }

        var sourceSheet = workbook.Worksheets[sourceSheetIndex];
        var destSheet = workbook.Worksheets[destSheetIdx];

        var sourceRangeObj = sourceSheet.Cells.CreateRange(sourceRange);
        var destRangeObj = destSheet.Cells.CreateRange(destCell);

        // Copy to destination
        destRangeObj.Copy(sourceRangeObj, new PasteOptions { PasteType = PasteType.All });

        // Clear source range
        for (int i = sourceRangeObj.FirstRow; i <= sourceRangeObj.FirstRow + sourceRangeObj.RowCount - 1; i++)
        {
            for (int j = sourceRangeObj.FirstColumn; j <= sourceRangeObj.FirstColumn + sourceRangeObj.ColumnCount - 1; j++)
            {
                sourceSheet.Cells[i, j].PutValue("");
            }
        }

        workbook.Save(path);
        return await Task.FromResult($"Range {sourceRange} moved to {destCell} in sheet {destSheetIdx}: {path}");
    }
}

