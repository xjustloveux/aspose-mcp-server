using System.Text.Json.Nodes;
using Aspose.Cells;

namespace AsposeMcpServer.Tools;

public class ExcelFormatCellsTool : IAsposeTool
{
    public string Description => "Format cells in Excel (font, color, borders, etc.)";

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
                description = "Cell range (e.g., 'A1:C5')"
            },
            fontName = new
            {
                type = "string",
                description = "Font name (optional, e.g., 'Arial', '微軟雅黑')"
            },
            bold = new
            {
                type = "boolean",
                description = "Bold font (optional)"
            },
            italic = new
            {
                type = "boolean",
                description = "Italic font (optional)"
            },
            fontSize = new
            {
                type = "number",
                description = "Font size (optional)"
            },
            fontColor = new
            {
                type = "string",
                description = "Font/text color (hex format like '#FF0000' or name like 'Red', optional)"
            },
            backgroundColor = new
            {
                type = "string",
                description = "Background color (hex format like '#FFFF00' or name like 'Yellow', optional)"
            }
        },
        required = new[] { "path", "range" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var sheetIndex = arguments?["sheetIndex"]?.GetValue<int>() ?? 0;
        var range = arguments?["range"]?.GetValue<string>() ?? throw new ArgumentException("range is required");
        var fontName = arguments?["fontName"]?.GetValue<string>();
        var bold = arguments?["bold"]?.GetValue<bool>();
        var italic = arguments?["italic"]?.GetValue<bool>();
        var fontSize = arguments?["fontSize"]?.GetValue<int>();
        var fontColor = arguments?["fontColor"]?.GetValue<string>();
        var backgroundColor = arguments?["backgroundColor"]?.GetValue<string>();

        using var workbook = new Workbook(path);
        if (sheetIndex < 0 || sheetIndex >= workbook.Worksheets.Count)
        {
            throw new ArgumentException($"sheetIndex must be between 0 and {workbook.Worksheets.Count - 1}");
        }

        var worksheet = workbook.Worksheets[sheetIndex];
        var cells = worksheet.Cells;
        var cellRange = cells.CreateRange(range);

        var style = workbook.CreateStyle();

        if (!string.IsNullOrEmpty(fontName))
        {
            style.Font.Name = fontName;
        }

        if (bold.HasValue)
        {
            style.Font.IsBold = bold.Value;
        }

        if (italic.HasValue)
        {
            style.Font.IsItalic = italic.Value;
        }

        if (fontSize.HasValue)
        {
            style.Font.Size = fontSize.Value;
        }

        if (!string.IsNullOrEmpty(fontColor))
        {
            try
            {
                var color = fontColor.StartsWith("#")
                    ? System.Drawing.ColorTranslator.FromHtml(fontColor)
                    : System.Drawing.Color.FromName(fontColor);
                style.Font.Color = color;
            }
            catch
            {
                // Ignore invalid color
            }
        }

        if (!string.IsNullOrEmpty(backgroundColor))
        {
            try
            {
                var color = backgroundColor.StartsWith("#")
                    ? System.Drawing.ColorTranslator.FromHtml(backgroundColor)
                    : System.Drawing.Color.FromName(backgroundColor);
                style.ForegroundColor = color;
                style.Pattern = BackgroundType.Solid;
            }
            catch
            {
                // Ignore invalid color
            }
        }

        cellRange.ApplyStyle(style, new StyleFlag { All = true });
        
        workbook.Save(path);

        return await Task.FromResult($"Cells formatted in range {range}: {path}");
    }
}

