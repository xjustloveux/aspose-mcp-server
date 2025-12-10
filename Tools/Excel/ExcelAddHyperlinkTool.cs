using System.Text.Json.Nodes;
using Aspose.Cells;

namespace AsposeMcpServer.Tools;

public class ExcelAddHyperlinkTool : IAsposeTool
{
    public string Description => "Add a hyperlink to a cell in an Excel worksheet";

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
            cell = new
            {
                type = "string",
                description = "Cell reference (e.g., 'A1')"
            },
            url = new
            {
                type = "string",
                description = "URL or file path for the hyperlink"
            },
            displayText = new
            {
                type = "string",
                description = "Display text for the hyperlink (optional, uses URL if not provided)"
            }
        },
        required = new[] { "path", "cell", "url" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var sheetIndex = arguments?["sheetIndex"]?.GetValue<int>() ?? 0;
        var cell = arguments?["cell"]?.GetValue<string>() ?? throw new ArgumentException("cell is required");
        var url = arguments?["url"]?.GetValue<string>() ?? throw new ArgumentException("url is required");
        var displayText = arguments?["displayText"]?.GetValue<string>();

        using var workbook = new Workbook(path);
        var worksheet = workbook.Worksheets[sheetIndex];
        var cellObj = worksheet.Cells[cell];

        if (!string.IsNullOrEmpty(displayText))
        {
            cellObj.PutValue(displayText);
        }

        worksheet.Hyperlinks.Add(cell, 1, 1, url);
        workbook.Save(path);

        return await Task.FromResult($"單元格 {cell} 已添加超連結: {url}");
    }
}

