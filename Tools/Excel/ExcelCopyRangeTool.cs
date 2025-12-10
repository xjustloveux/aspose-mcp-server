using System.Text.Json.Nodes;
using Aspose.Cells;

namespace AsposeMcpServer.Tools;

public class ExcelCopyRangeTool : IAsposeTool
{
    public string Description => "Copy range from source to destination in Excel";

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
            },
            copyOptions = new
            {
                type = "string",
                description = "Copy options: 'All', 'Values', 'Formats', 'Formulas' (optional, default: 'All')",
                @enum = new[] { "All", "Values", "Formats", "Formulas" }
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
        var copyOptions = arguments?["copyOptions"]?.GetValue<string>() ?? "All";

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

        var copyOptionsEnum = copyOptions switch
        {
            "Values" => PasteType.Values,
            "Formats" => PasteType.Formats,
            "Formulas" => PasteType.Formulas,
            _ => PasteType.All
        };

        destRangeObj.Copy(sourceRangeObj, new PasteOptions { PasteType = copyOptionsEnum });

        workbook.Save(path);
        return await Task.FromResult($"Range {sourceRange} copied to {destCell} in sheet {destSheetIdx}: {path}");
    }
}

