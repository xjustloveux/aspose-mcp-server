using System.Text.Json.Nodes;
using Aspose.Slides;
using Aspose.Slides.Charts;
using Aspose.Slides.Export;

namespace AsposeMcpServer.Tools;

public class PptEditChartTool : IAsposeTool
{
    public string Description => "Edit chart data, type, title, and format on a PowerPoint slide";

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
            },
            title = new
            {
                type = "string",
                description = "Chart title (optional)"
            },
            chartType = new
            {
                type = "string",
                description = "Chart type to change to (Column, Bar, Line, Pie, etc., optional)"
            },
            data = new
            {
                type = "object",
                description = "Chart data object with series and categories (optional)",
                properties = new
                {
                    categories = new
                    {
                        type = "array",
                        items = new { type = "string" },
                        description = "Category labels"
                    },
                    series = new
                    {
                        type = "array",
                        items = new
                        {
                            type = "object",
                            properties = new
                            {
                                name = new { type = "string" },
                                values = new
                                {
                                    type = "array",
                                    items = new { type = "number" }
                                }
                            }
                        },
                        description = "Series data"
                    }
                }
            }
        },
        required = new[] { "path", "slideIndex", "shapeIndex" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var slideIndex = arguments?["slideIndex"]?.GetValue<int>() ?? throw new ArgumentException("slideIndex is required");
        var shapeIndex = arguments?["shapeIndex"]?.GetValue<int>() ?? throw new ArgumentException("shapeIndex is required");
        var title = arguments?["title"]?.GetValue<string>();
        var chartTypeStr = arguments?["chartType"]?.GetValue<string>();
        var dataObj = arguments?["data"]?.AsObject();

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

        if (!string.IsNullOrEmpty(title))
        {
            chart.HasTitle = true;
            var chartTitle = chart.ChartTitle;
            if (chartTitle != null)
            {
                chartTitle.TextFrameForOverriding.Text = title;
            }
            else
            {
                chart.ChartTitle.AddTextFrameForOverriding(title);
            }
        }

        if (!string.IsNullOrEmpty(chartTypeStr))
        {
            var chartType = chartTypeStr.ToLower() switch
            {
                "column" => ChartType.ClusteredColumn,
                "bar" => ChartType.ClusteredBar,
                "line" => ChartType.Line,
                "pie" => ChartType.Pie,
                "area" => ChartType.Area,
                "scatter" => ChartType.ScatterWithSmoothLines,
                "doughnut" => ChartType.Doughnut,
                "bubble" => ChartType.Bubble,
                _ => chart.Type
            };
            chart.Type = chartType;
        }

        if (dataObj != null)
        {
            var chartData = chart.ChartData;
            chartData.Series.Clear();
            chartData.Categories.Clear();

            // Note: Chart data editing requires complex workbook operations with proper cell references
            // This functionality is complex and may require chart-specific implementations
            // For now, we'll provide a note that data editing should be done through ppt_update_chart_data
            throw new NotImplementedException("Chart data editing through edit_chart is complex. Please use ppt_update_chart_data for updating chart data.");
        }

        presentation.Save(path, SaveFormat.Pptx);
        return await Task.FromResult($"Chart updated on slide {slideIndex}, shape {shapeIndex}");
    }
}

