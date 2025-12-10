using System.Text.Json.Nodes;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeMcpServer.Tools;

public class ExcelAddChartTool : IAsposeTool
{
    public string Description => "Add a chart to an Excel worksheet";

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
            chartType = new
            {
                type = "string",
                description = "Chart type (Column, Bar, Line, Pie, etc.)"
            },
            dataRange = new
            {
                type = "string",
                description = "Data range for chart Y-axis (values, e.g., 'E2:E40')"
            },
            categoryAxisDataRange = new
            {
                type = "string",
                description = "Category axis (X-axis) data range (optional, e.g., 'A2:A40' for dates/labels)"
            },
            title = new
            {
                type = "string",
                description = "Chart title (optional)"
            }
        },
        required = new[] { "path", "chartType", "dataRange" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var sheetIndex = arguments?["sheetIndex"]?.GetValue<int>() ?? 0;
        var chartTypeStr = arguments?["chartType"]?.GetValue<string>() ?? throw new ArgumentException("chartType is required");
        var dataRange = arguments?["dataRange"]?.GetValue<string>() ?? throw new ArgumentException("dataRange is required");
        var categoryAxisDataRange = arguments?["categoryAxisDataRange"]?.GetValue<string>();
        var title = arguments?["title"]?.GetValue<string>();

        using var workbook = new Workbook(path);
        var worksheet = workbook.Worksheets[sheetIndex];

        var chartType = chartTypeStr.ToLower() switch
        {
            "column" => ChartType.Column,
            "bar" => ChartType.Bar,
            "line" => ChartType.Line,
            "pie" => ChartType.Pie,
            "area" => ChartType.Area,
            "scatter" => ChartType.Scatter,
            "doughnut" => ChartType.Doughnut,
            _ => ChartType.Column
        };

        int chartIndex = worksheet.Charts.Add(chartType, 5, 0, 20, 10);
        var chart = worksheet.Charts[chartIndex];
        
        // Clear existing series
        chart.NSeries.Clear();
        
        // Parse data range - support multiple ranges separated by comma (e.g., "E2:E40,G2:G40")
        var ranges = dataRange.Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
        
        if (!string.IsNullOrEmpty(categoryAxisDataRange))
        {
            // When category axis is provided, use SetChartDataRange with format: "categoryRange,valueRange"
            // But we need to be careful about how series are created
            string combinedRange = $"{categoryAxisDataRange},{dataRange}";
            chart.SetChartDataRange(combinedRange, true);
            
            // After SetChartDataRange, fix series configuration
            if (chart.NSeries.Count > 0)
            {
                // SetChartDataRange might create series with category data as the first series
                // We need to ensure only value ranges are used as series
                for (int i = 0; i < ranges.Length && i < chart.NSeries.Count; i++)
                {
                    var series = chart.NSeries[i];
                    // Ensure values are set correctly to the data range (Y-axis)
                    series.Values = ranges[i];
                    
                    // Clear the Name if it's set to category axis cell reference
                    // This prevents series name from being set to category axis cell
                    if (!string.IsNullOrEmpty(series.Name) && series.Name.Contains(categoryAxisDataRange.Split(':')[0]))
                    {
                        // Set a simple series name or clear it
                        series.Name = $"Series {i + 1}";
                    }
                    
                    // Try to set category data using reflection
                    try
                    {
                        var categoryDataProp = series.GetType().GetProperty("CategoryData");
                        if (categoryDataProp != null && categoryDataProp.CanWrite)
                        {
                            categoryDataProp.SetValue(series, categoryAxisDataRange);
                        }
                    }
                    catch
                    {
                        // If CategoryData property doesn't exist, continue
                    }
                }
                
                // Remove extra series if SetChartDataRange created too many
                if (chart.NSeries.Count > ranges.Length)
                {
                    while (chart.NSeries.Count > ranges.Length)
                    {
                        chart.NSeries.RemoveAt(chart.NSeries.Count - 1);
                    }
                }
            }
        }
        else
        {
            // No category axis: just add series with values data
            // Excel will automatically generate X-axis (1, 2, 3...)
            foreach (var range in ranges)
            {
                int seriesIndex = chart.NSeries.Add(range, true);
                var series = chart.NSeries[seriesIndex];
                series.Values = range;
            }
        }

        if (!string.IsNullOrEmpty(title))
        {
            chart.Title.Text = title;
        }

        workbook.Save(path);

        return await Task.FromResult($"Chart added to worksheet with data range: {dataRange}");
    }
}

