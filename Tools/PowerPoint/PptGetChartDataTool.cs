using System.Text.Json.Nodes;
using System.Text;
using Aspose.Slides;
using Aspose.Slides.Charts;

namespace AsposeMcpServer.Tools;

public class PptGetChartDataTool : IAsposeTool
{
    public string Description => "Get chart data, type, and information from a PowerPoint slide";

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
            shapeIndex = new
            {
                type = "number",
                description = "Shape index of the chart (0-based)"
            }
        },
        required = new[] { "path", "slideIndex", "shapeIndex" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var slideIndex = arguments?["slideIndex"]?.GetValue<int>() ?? throw new ArgumentException("slideIndex is required");
        var shapeIndex = arguments?["shapeIndex"]?.GetValue<int>() ?? throw new ArgumentException("shapeIndex is required");

        using var presentation = new Presentation(path);
        if (slideIndex < 0 || slideIndex >= presentation.Slides.Count)
        {
            throw new ArgumentException($"slideIndex must be between 0 and {presentation.Slides.Count - 1}");
        }

        var slide = presentation.Slides[slideIndex];
        if (shapeIndex < 0 || shapeIndex >= slide.Shapes.Count)
        {
            throw new ArgumentException($"shapeIndex must be between 0 and {slide.Shapes.Count - 1}");
        }

        var shape = slide.Shapes[shapeIndex];
        if (shape is not IChart chart)
        {
            throw new ArgumentException($"Shape at index {shapeIndex} is not a chart");
        }

        var sb = new StringBuilder();
        sb.AppendLine($"Chart Type: {chart.Type}");
        sb.AppendLine($"Has Title: {chart.HasTitle}");
        if (chart.HasTitle && chart.ChartTitle != null)
        {
            sb.AppendLine($"Title: {chart.ChartTitle.TextFrameForOverriding?.Text ?? ""}");
        }
        sb.AppendLine();

        var chartData = chart.ChartData;
        sb.AppendLine($"Categories ({chartData.Categories.Count}):");
        for (int i = 0; i < chartData.Categories.Count; i++)
        {
            var cat = chartData.Categories[i];
            sb.AppendLine($"  [{i}] {cat.Value}");
        }
        sb.AppendLine();

        sb.AppendLine($"Series ({chartData.Series.Count}):");
        for (int i = 0; i < chartData.Series.Count; i++)
        {
            var series = chartData.Series[i];
            sb.AppendLine($"  [{i}] {series.Name}");
            sb.AppendLine($"      Data Points: {series.DataPoints.Count}");
            for (int j = 0; j < Math.Min(series.DataPoints.Count, 10); j++)
            {
                var point = series.DataPoints[j];
                sb.AppendLine($"        [{j}] Value: {point.Value}");
            }
            if (series.DataPoints.Count > 10)
            {
                sb.AppendLine($"        ... ({series.DataPoints.Count - 10} more)");
            }
        }

        return await Task.FromResult(sb.ToString());
    }
}

