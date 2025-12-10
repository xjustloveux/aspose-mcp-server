using System.Text.Json.Nodes;
using Aspose.Cells;

namespace AsposeMcpServer.Tools;

public class ExcelBatchFormatCellsTool : IAsposeTool
{
    public string Description => "Batch format multiple cell ranges with same style in Excel";

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
            ranges = new
            {
                type = "array",
                items = new { type = "string" },
                description = "Array of cell ranges (e.g., ['A1:C5', 'E1:G5'])"
            },
            fontName = new
            {
                type = "string",
                description = "Font name (optional)"
            },
            fontSize = new
            {
                type = "number",
                description = "Font size (optional)"
            },
            bold = new
            {
                type = "boolean",
                description = "Bold (optional)"
            },
            backgroundColor = new
            {
                type = "string",
                description = "Background color (optional)"
            }
        },
        required = new[] { "path", "ranges" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var sheetIndex = arguments?["sheetIndex"]?.GetValue<int>() ?? 0;
        var rangesArray = arguments?["ranges"]?.AsArray() ?? throw new ArgumentException("ranges is required");
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

        var ranges = rangesArray.Select(r => r?.GetValue<string>()).Where(r => !string.IsNullOrEmpty(r)).ToList();
        foreach (var range in ranges)
        {
            var cellRange = cells.CreateRange(range!);
            cellRange.ApplyStyle(style, new StyleFlag { All = true });
        }

        workbook.Save(path);
        return await Task.FromResult($"Formatted {ranges.Count} range(s) in sheet {sheetIndex}: {path}");
    }
}

