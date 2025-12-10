using System.Text.Json.Nodes;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeMcpServer.Tools;

public class ExcelSetChartLegendTool : IAsposeTool
{
    public string Description => "Set chart legend position and visibility in Excel";

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
            visible = new
            {
                type = "boolean",
                description = "Legend visibility (optional)"
            },
            position = new
            {
                type = "string",
                description = "Legend position: 'Bottom', 'Top', 'Left', 'Right', 'TopRight' (optional)",
                @enum = new[] { "Bottom", "Top", "Left", "Right", "TopRight" }
            }
        },
        required = new[] { "path", "chartIndex" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var sheetIndex = arguments?["sheetIndex"]?.GetValue<int?>() ?? 0;
        var chartIndex = arguments?["chartIndex"]?.GetValue<int>() ?? throw new ArgumentException("chartIndex is required");
        var visible = arguments?["visible"]?.GetValue<bool?>();
        var position = arguments?["position"]?.GetValue<string>();

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

        if (visible.HasValue)
        {
            chart.ShowLegend = visible.Value;
        }

        if (!string.IsNullOrEmpty(position))
        {
            var legendPosition = position switch
            {
                "Bottom" => LegendPositionType.Bottom,
                "Top" => LegendPositionType.Top,
                "Left" => LegendPositionType.Left,
                "Right" => LegendPositionType.Right,
                "TopRight" => LegendPositionType.Right, // Use Right as fallback
                _ => LegendPositionType.Bottom
            };
            chart.Legend.Position = legendPosition;
        }

        workbook.Save(path);
        return await Task.FromResult($"Chart legend updated for chart #{chartIndex}: {path}");
    }
}

