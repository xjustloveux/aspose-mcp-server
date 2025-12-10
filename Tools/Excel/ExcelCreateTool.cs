using System.Text.Json.Nodes;
using Aspose.Cells;

namespace AsposeMcpServer.Tools;

public class ExcelCreateTool : IAsposeTool
{
    public string Description => "Create a new Excel workbook";

    public object InputSchema => new
    {
        type = "object",
        properties = new
        {
            path = new
            {
                type = "string",
                description = "Output file path"
            },
            sheetName = new
            {
                type = "string",
                description = "Initial sheet name (optional)"
            }
        },
        required = new[] { "path" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var sheetName = arguments?["sheetName"]?.GetValue<string>();

        using var workbook = new Workbook();
        
        if (!string.IsNullOrEmpty(sheetName))
        {
            workbook.Worksheets[0].Name = sheetName;
        }

        workbook.Save(path);
        return await Task.FromResult($"Excel workbook created successfully at: {path}");
    }
}

