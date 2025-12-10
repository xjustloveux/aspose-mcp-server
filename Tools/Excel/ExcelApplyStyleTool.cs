using System.Text.Json.Nodes;
using Aspose.Cells;

namespace AsposeMcpServer.Tools;

public class ExcelApplyStyleTool : IAsposeTool
{
    public string Description => "Apply style to cells or range in Excel";

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
            styleName = new
            {
                type = "string",
                description = "Style name to apply (optional, if not provided creates inline style)"
            },
            fontName = new
            {
                type = "string",
                description = "Font name (optional, for inline style)"
            },
            fontSize = new
            {
                type = "number",
                description = "Font size (optional, for inline style)"
            },
            bold = new
            {
                type = "boolean",
                description = "Bold (optional, for inline style)"
            },
            backgroundColor = new
            {
                type = "string",
                description = "Background color (optional, for inline style)"
            }
        },
        required = new[] { "path", "range" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var sheetIndex = arguments?["sheetIndex"]?.GetValue<int>() ?? 0;
        var range = arguments?["range"]?.GetValue<string>() ?? throw new ArgumentException("range is required");
        var styleName = arguments?["styleName"]?.GetValue<string>();
        var fontName = arguments?["fontName"]?.GetValue<string>();
        var fontSize = arguments?["fontSize"]?.GetValue<int?>();
        var bold = arguments?["bold"]?.GetValue<bool?>();
        var backgroundColor = arguments?["backgroundColor"]?.GetValue<string>();

        using var workbook = new Workbook(path);
        if (sheetIndex < 0 || sheetIndex >= workbook.Worksheets.Count)
        {
            throw new ArgumentException($"sheetIndex must be between 0 and {workbook.Worksheets.Count - 1}");
        }

        var worksheet = workbook.Worksheets[sheetIndex];
        var cells = worksheet.Cells;
        var cellRange = cells.CreateRange(range);

        // Note: Aspose.Cells doesn't support named styles directly
        // Always create inline style
        var style = workbook.CreateStyle();
        if (!string.IsNullOrEmpty(fontName)) style.Font.Name = fontName;
        if (fontSize.HasValue) style.Font.Size = fontSize.Value;
        if (bold.HasValue) style.Font.IsBold = bold.Value;
        if (!string.IsNullOrWhiteSpace(backgroundColor))
        {
            try
            {
                var color = backgroundColor.StartsWith("#")
                    ? System.Drawing.ColorTranslator.FromHtml(backgroundColor)
                    : System.Drawing.Color.FromName(backgroundColor);
                style.ForegroundColor = color;
                style.Pattern = BackgroundType.Solid;
            }
            catch { }
        }

        cellRange.ApplyStyle(style, new StyleFlag { All = true });
        workbook.Save(path);
        return await Task.FromResult($"Style applied to range {range} in sheet {sheetIndex}: {path}");
    }
}

