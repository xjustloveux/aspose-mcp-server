using System.Text.Json.Nodes;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeMcpServer.Tools;

public class ExcelAddPivotTableTool : IAsposeTool
{
    public string Description => "Add a pivot table to an Excel worksheet";

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
                description = "Source sheet index (0-based, optional, default: 0)"
            },
            sourceRange = new
            {
                type = "string",
                description = "Source data range (e.g., 'A1:D10')"
            },
            destCell = new
            {
                type = "string",
                description = "Destination cell for pivot table (e.g., 'F1')"
            }
        },
        required = new[] { "path", "sourceRange", "destCell" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var sheetIndex = arguments?["sheetIndex"]?.GetValue<int>() ?? 0;
        var sourceRange = arguments?["sourceRange"]?.GetValue<string>() ?? throw new ArgumentException("sourceRange is required");
        var destCell = arguments?["destCell"]?.GetValue<string>() ?? throw new ArgumentException("destCell is required");

        using var workbook = new Workbook(path);
        var worksheet = workbook.Worksheets[sheetIndex];

        var pivotTables = worksheet.PivotTables;
        int pivotIndex = pivotTables.Add($"={worksheet.Name}!{sourceRange}", destCell, "PivotTable1");
        var pivotTable = pivotTables[pivotIndex];

        // Add first field as row field and second field as data field by default
        pivotTable.AddFieldToArea(PivotFieldType.Row, 0);
        pivotTable.AddFieldToArea(PivotFieldType.Data, 1);
        
        pivotTable.CalculateData();

        workbook.Save(path);

        return await Task.FromResult($"Pivot table added to worksheet: {path}");
    }
}

