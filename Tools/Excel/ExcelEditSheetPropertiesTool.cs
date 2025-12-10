using System.Text.Json.Nodes;
using Aspose.Cells;

namespace AsposeMcpServer.Tools;

public class ExcelEditSheetPropertiesTool : IAsposeTool
{
    public string Description => "Edit worksheet properties (name, visibility, tab color, etc.)";

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
                description = "Sheet index (0-based)"
            },
            name = new
            {
                type = "string",
                description = "New sheet name (optional)"
            },
            isVisible = new
            {
                type = "boolean",
                description = "Sheet visibility (optional)"
            },
            tabColor = new
            {
                type = "string",
                description = "Tab color hex (e.g., #FF0000, optional)"
            },
            isSelected = new
            {
                type = "boolean",
                description = "Set as selected sheet (optional)"
            }
        },
        required = new[] { "path", "sheetIndex" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var sheetIndex = arguments?["sheetIndex"]?.GetValue<int>() ?? throw new ArgumentException("sheetIndex is required");
        var name = arguments?["name"]?.GetValue<string>();
        var isVisible = arguments?["isVisible"]?.GetValue<bool?>();
        var tabColor = arguments?["tabColor"]?.GetValue<string>();
        var isSelected = arguments?["isSelected"]?.GetValue<bool?>();

        using var workbook = new Workbook(path);
        if (sheetIndex < 0 || sheetIndex >= workbook.Worksheets.Count)
        {
            throw new ArgumentException($"sheetIndex must be between 0 and {workbook.Worksheets.Count - 1}");
        }

        var worksheet = workbook.Worksheets[sheetIndex];

        if (!string.IsNullOrEmpty(name))
        {
            worksheet.Name = name;
        }

        if (isVisible.HasValue)
        {
            worksheet.IsVisible = isVisible.Value;
        }

        if (!string.IsNullOrWhiteSpace(tabColor))
        {
            var color = System.Drawing.ColorTranslator.FromHtml(tabColor);
            worksheet.TabColor = color;
        }

        if (isSelected.HasValue && isSelected.Value)
        {
            workbook.Worksheets.ActiveSheetIndex = sheetIndex;
        }

        workbook.Save(path);
        return await Task.FromResult($"Sheet {sheetIndex} properties updated: {path}");
    }
}

