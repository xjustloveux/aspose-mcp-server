using System.Text.Json.Nodes;
using System.Linq;
using Aspose.Cells;

namespace AsposeMcpServer.Tools;

public class ExcelAddSheetTool : IAsposeTool
{
    public string Description => "Add a new worksheet to an Excel workbook";

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
            sheetName = new
            {
                type = "string",
                description = "Name of the new worksheet"
            },
            insertAt = new
            {
                type = "number",
                description = "Position to insert the sheet (0-based, optional, default: append at end)"
            }
        },
        required = new[] { "path", "sheetName" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var sheetName = arguments?["sheetName"]?.GetValue<string>()?.Trim() ?? throw new ArgumentException("sheetName is required");
        var insertAt = arguments?["insertAt"]?.GetValue<int?>();

        if (string.IsNullOrWhiteSpace(sheetName))
        {
            throw new ArgumentException("sheetName cannot be empty");
        }

        using var workbook = new Workbook(path);

        // Excel requires unique sheet names; avoid runtime exceptions by validating first.
        var duplicate = workbook.Worksheets.Any(ws => string.Equals(ws.Name, sheetName, StringComparison.OrdinalIgnoreCase));
        if (duplicate)
        {
            throw new ArgumentException($"Worksheet name '{sheetName}' already exists in the workbook");
        }

        Worksheet newSheet;
        if (insertAt.HasValue)
        {
            if (insertAt.Value < 0 || insertAt.Value > workbook.Worksheets.Count)
            {
                throw new ArgumentException($"insertAt must be between 0 and {workbook.Worksheets.Count}");
            }

            if (insertAt.Value == workbook.Worksheets.Count)
            {
                var addedIndex = workbook.Worksheets.Add();
                newSheet = workbook.Worksheets[addedIndex];
            }
            else
            {
                workbook.Worksheets.Insert(insertAt.Value, SheetType.Worksheet);
                newSheet = workbook.Worksheets[insertAt.Value];
            }
        }
        else
        {
            var addedIndex = workbook.Worksheets.Add();
            newSheet = workbook.Worksheets[addedIndex];
        }
        
        newSheet.Name = sheetName;
        workbook.Save(path);

        return await Task.FromResult($"Worksheet '{sheetName}' added: {path}");
    }
}
