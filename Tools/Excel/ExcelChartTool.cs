using System.Text.Json.Nodes;
using System.Text;
using Aspose.Cells;
using Aspose.Cells.Charts;
using AsposeMcpServer.Core;

namespace AsposeMcpServer.Tools;

/// <summary>
/// Unified tool for managing Excel charts (add, edit, delete, get, update data, set properties)
/// Merges: ExcelAddChartTool, ExcelEditChartTool, ExcelDeleteChartTool, ExcelGetChartsTool, 
/// ExcelUpdateChartDataTool, ExcelSetChartTitleTool, ExcelSetChartLegendTool
/// </summary>
public class ExcelChartTool : IAsposeTool
{
    public string Description => @"Manage Excel charts. Supports 6 operations: add, edit, delete, get, update_data, set_properties.

Usage examples:
- Add chart: excel_chart(operation='add', path='book.xlsx', chartType='Column', dataRange='A1:B10', position='A12')
- Edit chart: excel_chart(operation='edit', path='book.xlsx', chartIndex=0, chartType='Line')
- Delete chart: excel_chart(operation='delete', path='book.xlsx', chartIndex=0)
- Get charts: excel_chart(operation='get', path='book.xlsx')
- Update data: excel_chart(operation='update_data', path='book.xlsx', chartIndex=0, dataRange='A1:C10')
- Set properties: excel_chart(operation='set_properties', path='book.xlsx', chartIndex=0, title='Chart Title')";

    public object InputSchema => new
    {
        type = "object",
        properties = new
        {
            operation = new
            {
                type = "string",
                description = @"Operation to perform.
- 'add': Add a chart (required params: path, chartType, dataRange, position)
- 'edit': Edit chart properties (required params: path, chartIndex)
- 'delete': Delete a chart (required params: path, chartIndex)
- 'get': Get all charts (required params: path)
- 'update_data': Update chart data (required params: path, chartIndex, dataRange)
- 'set_properties': Set chart properties (required params: path, chartIndex)",
                @enum = new[] { "add", "edit", "delete", "get", "update_data", "set_properties" }
            },
            path = new
            {
                type = "string",
                description = "Excel file path (required for all operations)"
            },
            outputPath = new
            {
                type = "string",
                description = "Output file path (optional, for add/edit/delete/update_data/set_properties operations, defaults to input path)"
            },
            sheetIndex = new
            {
                type = "number",
                description = "Sheet index (0-based, optional, default: 0)"
            },
            chartIndex = new
            {
                type = "number",
                description = "Chart index (0-based, required for edit/delete/update_data/set_properties)"
            },
            // Add operation parameters
            chartType = new
            {
                type = "string",
                description = "Chart type: Column, Bar, Line, Pie, Area, Scatter, Doughnut, Radar, Bubble, Cylinder, Cone, Pyramid (required for add)",
                @enum = new[] { "Column", "Bar", "Line", "Pie", "Area", "Scatter", "Doughnut", "Radar", "Bubble", "Cylinder", "Cone", "Pyramid" }
            },
            dataRange = new
            {
                type = "string",
                description = "Data range for chart Y-axis (values, e.g., 'E2:E40')"
            },
            categoryAxisDataRange = new
            {
                type = "string",
                description = "Category axis (X-axis) data range (optional, e.g., 'A2:A40')"
            },
            title = new
            {
                type = "string",
                description = "Chart title (optional)"
            },
            topRow = new
            {
                type = "number",
                description = "Top row index for chart position (0-based, optional, default: auto-detect)"
            },
            leftColumn = new
            {
                type = "number",
                description = "Left column index for chart position (0-based, optional, default: 0)"
            },
            width = new
            {
                type = "number",
                description = "Chart width in columns (optional, default: 10)"
            },
            height = new
            {
                type = "number",
                description = "Chart height in rows (optional, default: 15)"
            },
            // Edit operation parameters
            showLegend = new
            {
                type = "boolean",
                description = "Show legend (optional, for edit/set_properties)"
            },
            legendPosition = new
            {
                type = "string",
                description = "Legend position (Bottom, Top, Left, Right, optional)"
            },
            // Set properties operation parameters
            removeTitle = new
            {
                type = "boolean",
                description = "Remove title (optional, for set_properties)"
            },
            legendVisible = new
            {
                type = "boolean",
                description = "Legend visibility (optional, for set_properties)"
            }
        },
        required = new[] { "operation", "path" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var operation = ArgumentHelper.GetString(arguments, "operation");
        var path = ArgumentHelper.GetAndValidatePath(arguments);
        var outputPath = ArgumentHelper.GetStringNullable(arguments, "outputPath");
        if (!string.IsNullOrEmpty(outputPath))
        {
            SecurityHelper.ValidateFilePath(outputPath, "outputPath");
        }
        var sheetIndex = ArgumentHelper.GetInt(arguments, "sheetIndex", 0);

        return operation.ToLower() switch
        {
            "add" => await AddChartAsync(arguments, path, sheetIndex),
            "edit" => await EditChartAsync(arguments, path, outputPath ?? path, sheetIndex),
            "delete" => await DeleteChartAsync(arguments, path, sheetIndex),
            "get" => await GetChartsAsync(arguments, path, sheetIndex),
            "update_data" => await UpdateChartDataAsync(arguments, path, outputPath ?? path, sheetIndex),
            "set_properties" => await SetChartPropertiesAsync(arguments, path, outputPath ?? path, sheetIndex),
            _ => throw new ArgumentException($"Unknown operation: {operation}")
        };
    }

    /// <summary>
    /// Adds a chart to the worksheet
    /// </summary>
    /// <param name="arguments">JSON arguments containing chartType, dataRange, optional categoryAxisDataRange, title, position</param>
    /// <param name="path">Excel file path</param>
    /// <param name="sheetIndex">Worksheet index (0-based)</param>
    /// <returns>Success message with chart index</returns>
    private async Task<string> AddChartAsync(JsonObject? arguments, string path, int sheetIndex)
    {
        var chartTypeStr = ArgumentHelper.GetString(arguments, "chartType");
        var dataRange = ArgumentHelper.GetString(arguments, "dataRange");
        var categoryAxisDataRange = ArgumentHelper.GetStringNullable(arguments, "categoryAxisDataRange");
        var title = ArgumentHelper.GetStringNullable(arguments, "title");
        var topRow = ArgumentHelper.GetIntNullable(arguments, "topRow");
        var leftColumn = ArgumentHelper.GetInt(arguments, "leftColumn", 0);
        var width = ArgumentHelper.GetInt(arguments, "width", 10);
        var height = ArgumentHelper.GetInt(arguments, "height", 15);

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
            "radar" => ChartType.Radar,
            "bubble" => ChartType.Bubble,
            "cylinder" => ChartType.Cylinder,
            "cone" => ChartType.Cone,
            "pyramid" => ChartType.Pyramid,
            _ => ChartType.Column
        };

        int chartTopRow;
        if (topRow.HasValue)
        {
            chartTopRow = topRow.Value;
        }
        else
        {
            var dataRangeObj = ExcelHelper.CreateRange(worksheet.Cells, dataRange.Split(',')[0].Trim());
            chartTopRow = dataRangeObj.FirstRow + dataRangeObj.RowCount + 2;
        }

        int chartIndex = worksheet.Charts.Add(chartType, chartTopRow, leftColumn, chartTopRow + height, leftColumn + width);
        var chart = worksheet.Charts[chartIndex];
        
        // Clear existing series
        chart.NSeries.Clear();
        
        // Parse data range - support multiple ranges separated by comma (e.g., "E2:E40,G2:G40")
        var ranges = dataRange.Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
        
        // Validate data range and category axis data range length match
        string? warningMessage = null;
        if (!string.IsNullOrEmpty(categoryAxisDataRange))
        {
            try
            {
                var dataRangeObj = ExcelHelper.CreateRange(worksheet.Cells, ranges[0]);
                var categoryRangeObj = ExcelHelper.CreateRange(worksheet.Cells, categoryAxisDataRange);
                
                if (dataRangeObj.RowCount != categoryRangeObj.RowCount)
                {
                    warningMessage = $"\n⚠️ Warning: Data range length ({dataRangeObj.RowCount} cells) does not match category axis length ({categoryRangeObj.RowCount} cells). Chart may display incorrectly.";
                }
            }
            catch
            {
                // Ignore validation errors, will be caught later
            }
        }
        
        // Add data series (Y-axis values) first
        foreach (var range in ranges)
        {
            int seriesIndex = chart.NSeries.Add(range, true);
            var series = chart.NSeries[seriesIndex];
            series.Values = range;
        }
        
        // Set category axis data (X-axis) if provided
        // Use NSeries.CategoryData property to set category data for all series
        if (!string.IsNullOrEmpty(categoryAxisDataRange) && chart.NSeries.Count > 0)
        {
            try
            {
                // Try to set category data using NSeries collection property
                var categoryDataProp = chart.NSeries.GetType().GetProperty("CategoryData");
                if (categoryDataProp != null && categoryDataProp.CanWrite)
                {
                    categoryDataProp.SetValue(chart.NSeries, categoryAxisDataRange);
                }
                else
                {
                    // Fallback: Set category data for each series individually
                    foreach (Aspose.Cells.Charts.Series series in chart.NSeries)
                    {
                        try
                        {
                            var seriesCategoryDataProp = series.GetType().GetProperty("CategoryData");
                            if (seriesCategoryDataProp != null && seriesCategoryDataProp.CanWrite)
                            {
                                seriesCategoryDataProp.SetValue(series, categoryAxisDataRange);
                            }
                        }
                        catch { }
                    }
                }
            }
            catch
            {
                // Final fallback: Use SetChartDataRange but fix the series afterwards
                chart.NSeries.Clear();
                string combinedRange = $"{categoryAxisDataRange},{dataRange}";
                chart.SetChartDataRange(combinedRange, true);
                
                // Remove extra series if SetChartDataRange created too many
                while (chart.NSeries.Count > ranges.Length)
                {
                    chart.NSeries.RemoveAt(chart.NSeries.Count - 1);
                }
                
                // Ensure all series have correct values
                for (int i = 0; i < ranges.Length && i < chart.NSeries.Count; i++)
                {
                    chart.NSeries[i].Values = ranges[i];
                }
            }
        }

        if (!string.IsNullOrEmpty(title))
        {
            chart.Title.Text = title;
        }

        workbook.CalculateFormula();
        var outputPath = ArgumentHelper.GetAndValidateOutputPath(arguments, path);
        workbook.Save(outputPath);

        var result = $"Chart added to worksheet with data range: {dataRange}";
        if (!string.IsNullOrEmpty(categoryAxisDataRange))
        {
            result += $", X-axis: {categoryAxisDataRange}";
        }
        result += $"\nOutput: {outputPath}";
        if (!string.IsNullOrEmpty(warningMessage))
        {
            result += warningMessage;
        }
        return await Task.FromResult(result);
    }

    /// <summary>
    /// Edits chart properties
    /// </summary>
    /// <param name="arguments">JSON arguments containing chartIndex and various chart properties</param>
    /// <param name="path">Excel file path</param>
    /// <param name="outputPath">Output file path</param>
    /// <param name="sheetIndex">Worksheet index (0-based)</param>
    /// <returns>Success message</returns>
    private async Task<string> EditChartAsync(JsonObject? arguments, string path, string outputPath, int sheetIndex)
    {
        var chartIndex = ArgumentHelper.GetInt(arguments, "chartIndex");
        var title = ArgumentHelper.GetStringNullable(arguments, "title");
        var dataRange = ArgumentHelper.GetStringNullable(arguments, "dataRange");
        var categoryAxisDataRange = ArgumentHelper.GetStringNullable(arguments, "categoryAxisDataRange");
        var chartTypeStr = ArgumentHelper.GetStringNullable(arguments, "chartType");
        var showLegend = ArgumentHelper.GetBoolNullable(arguments, "showLegend");
        var legendPosition = ArgumentHelper.GetStringNullable(arguments, "legendPosition");

        using var workbook = new Workbook(path);
        
        var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);
        var charts = worksheet.Charts;
        
        if (chartIndex < 0 || chartIndex >= charts.Count)
        {
            throw new ArgumentException($"Chart index {chartIndex} is out of range (worksheet has {charts.Count} charts)");
        }

        var chart = charts[chartIndex];
        var changes = new List<string>();

        if (!string.IsNullOrEmpty(title))
        {
            chart.Title.Text = title;
            changes.Add($"Title: {title}");
        }

        if (!string.IsNullOrEmpty(dataRange))
        {
            chart.NSeries.Clear();
            
            var ranges = dataRange.Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
            
            // Add data series (Y-axis values)
            foreach (var range in ranges)
            {
                int seriesIndex = chart.NSeries.Add(range, true);
                var series = chart.NSeries[seriesIndex];
                series.Values = range;
            }
            
            // Set category axis data (X-axis) if provided
            if (!string.IsNullOrEmpty(categoryAxisDataRange) && chart.NSeries.Count > 0)
            {
                try
                {
                    // Try to set category data using NSeries collection property
                    var categoryDataProp = chart.NSeries.GetType().GetProperty("CategoryData");
                    if (categoryDataProp != null && categoryDataProp.CanWrite)
                    {
                        categoryDataProp.SetValue(chart.NSeries, categoryAxisDataRange);
                    }
                    else
                    {
                        // Fallback: Use SetChartDataRange but fix the series afterwards
                        chart.NSeries.Clear();
                        string combinedRange = $"{categoryAxisDataRange},{dataRange}";
                        chart.SetChartDataRange(combinedRange, true);
                        
                        // Remove extra series if SetChartDataRange created too many
                        while (chart.NSeries.Count > ranges.Length)
                        {
                            chart.NSeries.RemoveAt(chart.NSeries.Count - 1);
                        }
                        
                        // Ensure all series have correct values
                        for (int i = 0; i < ranges.Length && i < chart.NSeries.Count; i++)
                        {
                            chart.NSeries[i].Values = ranges[i];
                        }
                    }
                }
                catch
                {
                    // Final fallback: Use SetChartDataRange and clean up
                    chart.NSeries.Clear();
                    string combinedRange = $"{categoryAxisDataRange},{dataRange}";
                    chart.SetChartDataRange(combinedRange, true);
                    
                    // Remove extra series
                    while (chart.NSeries.Count > ranges.Length)
                    {
                        chart.NSeries.RemoveAt(chart.NSeries.Count - 1);
                    }
                    
                    // Ensure all series have correct values
                    for (int i = 0; i < ranges.Length && i < chart.NSeries.Count; i++)
                    {
                        chart.NSeries[i].Values = ranges[i];
                    }
                }
            }
            
            var rangeInfo = dataRange;
            if (!string.IsNullOrEmpty(categoryAxisDataRange))
            {
                rangeInfo += $", X-axis: {categoryAxisDataRange}";
            }
            changes.Add($"Data range: {rangeInfo}");
        }

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
                "radar" => ChartType.Radar,
                "bubble" => ChartType.Bubble,
                "cylinder" => ChartType.Cylinder,
                "cone" => ChartType.Cone,
                "pyramid" => ChartType.Pyramid,
                _ => chart.Type
            };
            chart.Type = chartType;
            changes.Add($"Chart type: {chartTypeStr}");
        }

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
            changes.Add($"Legend: {(showLegend.Value ? "show" : "hide")}");
        }

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
            changes.Add($"Legend position: {legendPosition}");
        }

        workbook.Save(outputPath);

        var result = $"Successfully edited chart #{chartIndex}\n";
        if (changes.Count > 0)
        {
            result += "Changes:\n";
            foreach (var change in changes)
            {
                result += $"  - {change}\n";
            }
        }
        else
        {
            result += "No changes.\n";
        }
        result += $"Output: {outputPath}";

        return await Task.FromResult(result);
    }

    /// <summary>
    /// Deletes a chart from the worksheet
    /// </summary>
    /// <param name="arguments">JSON arguments containing chartIndex</param>
    /// <param name="path">Excel file path</param>
    /// <param name="sheetIndex">Worksheet index (0-based)</param>
    /// <returns>Success message</returns>
    private async Task<string> DeleteChartAsync(JsonObject? arguments, string path, int sheetIndex)
    {
        var chartIndex = ArgumentHelper.GetInt(arguments, "chartIndex");

        using var workbook = new Workbook(path);
        
        var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);
        var charts = worksheet.Charts;
        
        if (chartIndex < 0 || chartIndex >= charts.Count)
        {
            throw new ArgumentException($"Chart index {chartIndex} is out of range (worksheet has {charts.Count} charts)");
        }

        var chart = charts[chartIndex];
        var chartName = chart.Name ?? $"Chart {chartIndex}";
        
        charts.RemoveAt(chartIndex);
        var outputPath = ArgumentHelper.GetAndValidateOutputPath(arguments, path);
        workbook.Save(outputPath);

        var remainingCount = charts.Count;
        
        return await Task.FromResult($"Successfully deleted chart #{chartIndex} ({chartName})\nRemaining charts in worksheet: {remainingCount}\nOutput: {outputPath}");
    }

    /// <summary>
    /// Gets all charts from the worksheet
    /// </summary>
    /// <param name="arguments">JSON arguments (no specific parameters required)</param>
    /// <param name="path">Excel file path</param>
    /// <param name="sheetIndex">Worksheet index (0-based)</param>
    /// <returns>Formatted string with all charts</returns>
    private async Task<string> GetChartsAsync(JsonObject? arguments, string path, int sheetIndex)
    {
        using var workbook = new Workbook(path);
        
        var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);
        var charts = worksheet.Charts;
        var result = new StringBuilder();

        result.AppendLine($"=== Chart information for worksheet '{worksheet.Name}' ===\n");
        result.AppendLine($"Total charts: {charts.Count}\n");

        if (charts.Count == 0)
        {
            result.AppendLine("No charts found");
            return await Task.FromResult(result.ToString());
        }

        for (int i = 0; i < charts.Count; i++)
        {
            var chart = charts[i];
            result.AppendLine($"[Chart {i}]");
            result.AppendLine($"Name: {chart.Name ?? "(no name)"}");
            result.AppendLine($"Type: {chart.Type}");
            result.AppendLine($"Location: rows {chart.ChartObject.UpperLeftRow} - {chart.ChartObject.LowerRightRow}, columns {chart.ChartObject.UpperLeftColumn} - {chart.ChartObject.LowerRightColumn}");
            result.AppendLine($"Width: {chart.ChartObject.Width}");
            result.AppendLine($"Height: {chart.ChartObject.Height}");
            
            if (chart.NSeries != null && chart.NSeries.Count > 0)
            {
                result.AppendLine($"Data series count: {chart.NSeries.Count}");
                for (int j = 0; j < chart.NSeries.Count && j < 5; j++)
                {
                    var series = chart.NSeries[j];
                    var seriesName = series.Name ?? "(no name)";
                    var valuesRange = series.Values ?? "";
                    
                    result.AppendLine($"  Series {j}: {seriesName}");
                    if (!string.IsNullOrEmpty(valuesRange))
                    {
                        result.AppendLine($"    Data range (Y-axis): {valuesRange}");
                    }
                    
                    string? categoryData = null;
                    
                    try
                    {
                        var chartDataRange = chart.GetChartDataRange();
                        if (!string.IsNullOrEmpty(chartDataRange))
                        {
                            var parts = chartDataRange.Split(',');
                            if (parts.Length >= 2)
                            {
                                var potentialCategory = parts[0].Trim();
                                if (potentialCategory.Contains(":") || potentialCategory.Contains("$"))
                                {
                                    categoryData = potentialCategory;
                                }
                            }
                        }
                    }
                    catch { }
                    
                    if (string.IsNullOrEmpty(categoryData) || (!categoryData.Contains(":") && categoryData.Contains("$")))
                    {
                        try
                        {
                            var categoryDataProp = series.GetType().GetProperty("CategoryData");
                            if (categoryDataProp != null)
                            {
                                var catData = categoryDataProp.GetValue(series);
                                if (catData != null)
                                {
                                    var catDataStr = catData.ToString();
                                    if (!string.IsNullOrEmpty(catDataStr))
                                    {
                                        if (catDataStr.Contains(":"))
                                        {
                                            categoryData = catDataStr;
                                        }
                                        else if (string.IsNullOrEmpty(categoryData) && 
                                                 (catDataStr.Contains("$") || catDataStr.StartsWith("=")))
                                        {
                                            categoryData = catDataStr;
                                        }
                                    }
                                }
                            }
                        }
                        catch { }
                    }
                    
                    if (!string.IsNullOrEmpty(categoryData))
                    {
                        result.AppendLine($"    X-axis data: {categoryData}");
                    }
                }
            }
            
            if (chart.Title != null)
            {
                result.AppendLine($"Title: {chart.Title.Text}");
            }
            
            if (chart.Legend != null)
            {
                result.AppendLine($"Legend: enabled");
            }
            result.AppendLine();
        }

        return await Task.FromResult(result.ToString());
    }

    /// <summary>
    /// Updates chart data range
    /// </summary>
    /// <param name="arguments">JSON arguments containing chartIndex, dataRange, optional categoryAxisDataRange</param>
    /// <param name="path">Excel file path</param>
    /// <param name="outputPath">Output file path</param>
    /// <param name="sheetIndex">Worksheet index (0-based)</param>
    /// <returns>Success message</returns>
    private async Task<string> UpdateChartDataAsync(JsonObject? arguments, string path, string outputPath, int sheetIndex)
    {
        var chartIndex = ArgumentHelper.GetInt(arguments, "chartIndex");
        var dataRange = ArgumentHelper.GetString(arguments, "dataRange");

        using var workbook = new Workbook(path);
        
        var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);
        var charts = worksheet.Charts;
        
        if (chartIndex < 0 || chartIndex >= charts.Count)
        {
            throw new ArgumentException($"Chart index {chartIndex} is out of range (worksheet has {charts.Count} charts)");
        }

        var chart = charts[chartIndex];
        
        chart.NSeries.Clear();
        
        var ranges = dataRange.Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
        
        foreach (var range in ranges)
        {
            int seriesIndex = chart.NSeries.Add(range, true);
            var series = chart.NSeries[seriesIndex];
            series.Values = range;
        }
        
        if (chart.NSeries.Count == 0)
        {
            chart.SetChartDataRange(dataRange, true);
        }

        workbook.Save(outputPath);

        return await Task.FromResult($"Successfully updated data source for chart #{chartIndex}\nNew data range: {dataRange}\nOutput: {outputPath}");
    }

    /// <summary>
    /// Sets chart properties
    /// </summary>
    /// <param name="arguments">JSON arguments containing chartIndex and various chart properties</param>
    /// <param name="path">Excel file path</param>
    /// <param name="outputPath">Output file path</param>
    /// <param name="sheetIndex">Worksheet index (0-based)</param>
    /// <returns>Success message</returns>
    private async Task<string> SetChartPropertiesAsync(JsonObject? arguments, string path, string outputPath, int sheetIndex)
    {
        var chartIndex = ArgumentHelper.GetInt(arguments, "chartIndex");
        var title = ArgumentHelper.GetStringNullable(arguments, "title");
        var removeTitle = ArgumentHelper.GetBool(arguments, "removeTitle", false);
        var legendVisible = ArgumentHelper.GetBoolNullable(arguments, "legendVisible");
        var legendPosition = ArgumentHelper.GetStringNullable(arguments, "legendPosition");

        using var workbook = new Workbook(path);
        var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);
        PowerPointHelper.ValidateCollectionIndex(chartIndex, worksheet.Charts, "chart");

        var chart = worksheet.Charts[chartIndex];
        var changes = new List<string>();

        // Handle title
        if (removeTitle)
        {
            chart.Title.Text = "";
            changes.Add("Title removed");
        }
        else if (!string.IsNullOrEmpty(title))
        {
            chart.Title.Text = title;
            changes.Add($"Title: {title}");
        }

        // Handle legend visibility
        if (legendVisible.HasValue)
        {
            chart.ShowLegend = legendVisible.Value;
            changes.Add($"Legend: {(legendVisible.Value ? "show" : "hide")}");
        }

        // Handle legend position
        if (!string.IsNullOrEmpty(legendPosition) && chart.Legend != null)
        {
            var position = legendPosition switch
            {
                "Bottom" => LegendPositionType.Bottom,
                "Top" => LegendPositionType.Top,
                "Left" => LegendPositionType.Left,
                "Right" => LegendPositionType.Right,
                "TopRight" => LegendPositionType.Right,
                _ => chart.Legend.Position
            };
            chart.Legend.Position = position;
            changes.Add($"Legend position: {legendPosition}");
        }

        workbook.Save(outputPath);
        
        var result = changes.Count > 0 
            ? $"Chart properties updated: {string.Join(", ", changes)}\nOutput: {outputPath}"
            : $"No changes\nOutput: {outputPath}";
        
        return await Task.FromResult(result);
    }

}

