using System.Text.Json.Nodes;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeMcpServer.Tools;

public class ExcelSetPageSetupTool : IAsposeTool
{
    public string Description => "Set page setup options for an Excel worksheet (margins, orientation, paper size, headers/footers)";

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
            orientation = new
            {
                type = "string",
                description = "Page orientation: 'Portrait' or 'Landscape'",
                @enum = new[] { "Portrait", "Landscape" }
            },
            paperSize = new
            {
                type = "string",
                description = "Paper size (e.g., 'A4', 'Letter', 'Legal')"
            },
            leftMargin = new
            {
                type = "number",
                description = "Left margin in inches (optional)"
            },
            rightMargin = new
            {
                type = "number",
                description = "Right margin in inches (optional)"
            },
            topMargin = new
            {
                type = "number",
                description = "Top margin in inches (optional)"
            },
            bottomMargin = new
            {
                type = "number",
                description = "Bottom margin in inches (optional)"
            },
            header = new
            {
                type = "string",
                description = "Header text (optional)"
            },
            footer = new
            {
                type = "string",
                description = "Footer text (optional)"
            }
        },
        required = new[] { "path" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var sheetIndex = arguments?["sheetIndex"]?.GetValue<int>() ?? 0;
        var orientation = arguments?["orientation"]?.GetValue<string>();
        var paperSize = arguments?["paperSize"]?.GetValue<string>();
        var leftMargin = arguments?["leftMargin"]?.GetValue<double?>();
        var rightMargin = arguments?["rightMargin"]?.GetValue<double?>();
        var topMargin = arguments?["topMargin"]?.GetValue<double?>();
        var bottomMargin = arguments?["bottomMargin"]?.GetValue<double?>();
        var header = arguments?["header"]?.GetValue<string>();
        var footer = arguments?["footer"]?.GetValue<string>();

        using var workbook = new Workbook(path);
        var worksheet = workbook.Worksheets[sheetIndex];
        var pageSetup = worksheet.PageSetup;

        if (!string.IsNullOrEmpty(orientation))
        {
            pageSetup.Orientation = orientation == "Landscape" ? PageOrientationType.Landscape : PageOrientationType.Portrait;
        }

        if (!string.IsNullOrEmpty(paperSize))
        {
            var size = paperSize.ToUpper() switch
            {
                "A4" => PaperSizeType.PaperA4,
                "LETTER" => PaperSizeType.PaperLetter,
                "LEGAL" => PaperSizeType.PaperLegal,
                "A3" => PaperSizeType.PaperA3,
                "A5" => PaperSizeType.PaperA5,
                "B4" => PaperSizeType.PaperB4,
                "B5" => PaperSizeType.PaperB5,
                _ => PaperSizeType.PaperA4
            };
            pageSetup.PaperSize = size;
        }

        if (leftMargin.HasValue) pageSetup.LeftMargin = leftMargin.Value;
        if (rightMargin.HasValue) pageSetup.RightMargin = rightMargin.Value;
        if (topMargin.HasValue) pageSetup.TopMargin = topMargin.Value;
        if (bottomMargin.HasValue) pageSetup.BottomMargin = bottomMargin.Value;

        if (!string.IsNullOrEmpty(header))
        {
            pageSetup.SetHeader(0, header);
        }

        if (!string.IsNullOrEmpty(footer))
        {
            pageSetup.SetFooter(0, footer);
        }

        workbook.Save(path);

        var changes = new List<string>();
        if (!string.IsNullOrEmpty(orientation)) changes.Add($"方向: {orientation}");
        if (!string.IsNullOrEmpty(paperSize)) changes.Add($"紙張大小: {paperSize}");
        if (leftMargin.HasValue || rightMargin.HasValue || topMargin.HasValue || bottomMargin.HasValue)
            changes.Add("邊距已設定");
        if (!string.IsNullOrEmpty(header)) changes.Add("頁首已設定");
        if (!string.IsNullOrEmpty(footer)) changes.Add("頁尾已設定");

        return await Task.FromResult($"頁面設定已更新: {string.Join(", ", changes)}");
    }
}

