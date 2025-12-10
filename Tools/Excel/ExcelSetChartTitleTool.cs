using System.Text.Json.Nodes;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeMcpServer.Tools;

public class ExcelSetChartTitleTool : IAsposeTool
{
    public string Description => "Set chart title in Excel";

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
            chartIndex = new
            {
                type = "number",
                description = "Chart index (0-based)"
            },
            title = new
            {
                type = "string",
                description = "Chart title text"
            },
            removeTitle = new
            {
                type = "boolean",
                description = "Remove title (optional, default: false)"
            }
        },
        required = new[] { "path", "chartIndex" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var sheetIndex = arguments?["sheetIndex"]?.GetValue<int?>() ?? 0;
        var chartIndex = arguments?["chartIndex"]?.GetValue<int>() ?? throw new ArgumentException("chartIndex is required");
        var title = arguments?["title"]?.GetValue<string>();
        var removeTitle = arguments?["removeTitle"]?.GetValue<bool?>() ?? false;

        using var workbook = new Workbook(path);
        if (sheetIndex < 0 || sheetIndex >= workbook.Worksheets.Count)
        {
            throw new ArgumentException($"sheetIndex must be between 0 and {workbook.Worksheets.Count - 1}");
        }

        var worksheet = workbook.Worksheets[sheetIndex];
        if (chartIndex < 0 || chartIndex >= worksheet.Charts.Count)
        {
            throw new ArgumentException($"chartIndex must be between 0 and {worksheet.Charts.Count - 1}");
        }

        var chart = worksheet.Charts[chartIndex];

        if (removeTitle)
        {
            chart.Title.Text = "";
        }
        else if (!string.IsNullOrEmpty(title))
        {
            chart.Title.Text = title;
        }
        else
        {
            throw new ArgumentException("Either title or removeTitle must be provided");
        }

        workbook.Save(path);
        return await Task.FromResult(removeTitle
            ? $"Chart title removed from chart #{chartIndex}: {path}"
            : $"Chart title set to '{title}' for chart #{chartIndex}: {path}");
    }
}

