using System.Text.Json.Nodes;
using Aspose.Slides;
using Aspose.Slides.Charts;
using Aspose.Slides.Export;

namespace AsposeMcpServer.Tools;

public class PptAddChartTool : IAsposeTool
{
    public string Description => "Add a chart to a PowerPoint slide";

    public object InputSchema => new
    {
        type = "object",
        properties = new
        {
            path = new
            {
                type = "string",
                description = "Presentation file path"
            },
            slideIndex = new
            {
                type = "number",
                description = "Slide index (0-based)"
            },
            chartType = new
            {
                type = "string",
                description = "Chart type (Column, Bar, Line, Pie, etc.)"
            },
            title = new
            {
                type = "string",
                description = "Chart title (optional)"
            }
        },
        required = new[] { "path", "slideIndex", "chartType" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var slideIndex = arguments?["slideIndex"]?.GetValue<int>() ?? throw new ArgumentException("slideIndex is required");
        var chartTypeStr = arguments?["chartType"]?.GetValue<string>() ?? throw new ArgumentException("chartType is required");
        var title = arguments?["title"]?.GetValue<string>();

        using var presentation = new Presentation(path);
        if (slideIndex < 0 || slideIndex >= presentation.Slides.Count)
        {
            throw new ArgumentException($"slideIndex must be between 0 and {presentation.Slides.Count - 1}");
        }

        var slide = presentation.Slides[slideIndex];

        var chartType = chartTypeStr.ToLower() switch
        {
            "column" => ChartType.ClusteredColumn,
            "bar" => ChartType.ClusteredBar,
            "line" => ChartType.Line,
            "pie" => ChartType.Pie,
            "area" => ChartType.Area,
            "scatter" => ChartType.ScatterWithSmoothLines,
            _ => ChartType.ClusteredColumn
        };

        var chart = slide.Shapes.AddChart(chartType, 50, 50, 500, 400);

        if (!string.IsNullOrEmpty(title))
        {
            chart.HasTitle = true;
            chart.ChartTitle.AddTextFrameForOverriding(title);
        }

        presentation.Save(path, SaveFormat.Pptx);

        return await Task.FromResult($"Chart added to slide {slideIndex}: {path}");
    }
}

