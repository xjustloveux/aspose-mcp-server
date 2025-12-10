using System.Text.Json.Nodes;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeMcpServer.Tools;

public class ExcelEditChartTool : IAsposeTool
{
    public string Description => "Edit an existing chart in an Excel worksheet";

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
            outputPath = new
            {
                type = "string",
                description = "Output file path (if not provided, overwrites input)"
            },
            sheetIndex = new
            {
                type = "number",
                description = "Sheet index (0-based, optional, default: 0)"
            },
            chartIndex = new
            {
                type = "number",
                description = "Chart index to edit (0-based)"
            },
            title = new
            {
                type = "string",
                description = "New chart title (optional)"
            },
            dataRange = new
            {
                type = "string",
                description = "New data range for chart (e.g., 'A1:B10', optional)"
            },
            chartType = new
            {
                type = "string",
                description = "New chart type (Column, Bar, Line, Pie, etc., optional)"
            },
            showLegend = new
            {
                type = "boolean",
                description = "Show legend (optional)"
            },
            legendPosition = new
            {
                type = "string",
                description = "Legend position (Bottom, Top, Left, Right, optional)"
            }
        },
        required = new[] { "path", "chartIndex" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var outputPath = arguments?["outputPath"]?.GetValue<string>() ?? path;
        var sheetIndex = arguments?["sheetIndex"]?.GetValue<int>() ?? 0;
        var chartIndex = arguments?["chartIndex"]?.GetValue<int>() ?? throw new ArgumentException("chartIndex is required");
        var title = arguments?["title"]?.GetValue<string>();
        var dataRange = arguments?["dataRange"]?.GetValue<string>();
        var chartTypeStr = arguments?["chartType"]?.GetValue<string>();
        var showLegend = arguments?["showLegend"]?.GetValue<bool?>();
        var legendPosition = arguments?["legendPosition"]?.GetValue<string>();

        using var workbook = new Workbook(path);
        
        if (sheetIndex < 0 || sheetIndex >= workbook.Worksheets.Count)
        {
            throw new ArgumentException($"工作表索引 {sheetIndex} 超出範圍 (共有 {workbook.Worksheets.Count} 個工作表)");
        }

        var worksheet = workbook.Worksheets[sheetIndex];
        var charts = worksheet.Charts;
        
        if (chartIndex < 0 || chartIndex >= charts.Count)
        {
            throw new ArgumentException($"圖表索引 {chartIndex} 超出範圍 (工作表共有 {charts.Count} 個圖表)");
        }

        var chart = charts[chartIndex];
        var changes = new List<string>();

        // Update title
        if (!string.IsNullOrEmpty(title))
        {
            chart.Title.Text = title;
            changes.Add($"標題: {title}");
        }

        // Update data range
        if (!string.IsNullOrEmpty(dataRange))
        {
            chart.SetChartDataRange(dataRange, true);
            changes.Add($"數據範圍: {dataRange}");
        }

        // Update chart type
        if (!string.IsNullOrEmpty(chartTypeStr))
        {
            var chartType = chartTypeStr.ToLower() switch
            {
                "column" => ChartType.Column,
                "bar" => ChartType.Bar,
                "line" => ChartType.Line,
                "pie" => ChartType.Pie,
                "area" => ChartType.Area,
                "scatter" => ChartType.Scatter,
                "doughnut" => ChartType.Doughnut,
                _ => chart.Type
            };
            chart.Type = chartType;
            changes.Add($"圖表類型: {chartTypeStr}");
        }

        // Update legend
        if (showLegend.HasValue)
        {
            if (showLegend.Value && chart.Legend == null)
            {
                chart.ShowLegend = true;
            }
            else if (!showLegend.Value && chart.Legend != null)
            {
                chart.ShowLegend = false;
            }
            changes.Add($"圖例: {(showLegend.Value ? "顯示" : "隱藏")}");
        }

        // Update legend position
        if (!string.IsNullOrEmpty(legendPosition) && chart.Legend != null)
        {
            var position = legendPosition.ToLower() switch
            {
                "bottom" => LegendPositionType.Bottom,
                "top" => LegendPositionType.Top,
                "left" => LegendPositionType.Left,
                "right" => LegendPositionType.Right,
                _ => chart.Legend.Position
            };
            chart.Legend.Position = position;
            changes.Add($"圖例位置: {legendPosition}");
        }

        workbook.Save(outputPath);

        var result = $"成功編輯圖表 #{chartIndex}\n";
        if (changes.Count > 0)
        {
            result += "變更:\n";
            foreach (var change in changes)
            {
                result += $"  - {change}\n";
            }
        }
        else
        {
            result += "無變更。\n";
        }
        result += $"輸出: {outputPath}";

        return await Task.FromResult(result);
    }
}

