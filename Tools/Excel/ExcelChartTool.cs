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
    public string Description => "Manage Excel charts: add, edit, delete, get, update data, or set properties (title, legend)";

    public object InputSchema => new
    {
        type = "object",
        properties = new
        {
            operation = new
            {
                type = "string",
                description = "Operation to perform: 'add', 'edit', 'delete', 'get', 'update_data', 'set_properties'",
                @enum = new[] { "add", "edit", "delete", "get", "update_data", "set_properties" }
            },
            path = new
            {
                type = "string",
                description = "Excel file path"
            },
            outputPath = new
            {
                type = "string",
                description = "Output file path (optional, for edit/update_data operations, defaults to input path)"
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
                description = "Chart type (Column, Bar, Line, Pie, etc., required for add)"
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
        var operation = arguments?["operation"]?.GetValue<string>() ?? throw new ArgumentException("operation is required");
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        SecurityHelper.ValidateFilePath(path, "path");
        var outputPath = arguments?["outputPath"]?.GetValue<string>();
        if (!string.IsNullOrEmpty(outputPath))
        {
            SecurityHelper.ValidateFilePath(outputPath, "outputPath");
        }
        var sheetIndex = arguments?["sheetIndex"]?.GetValue<int>() ?? 0;

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

    private async Task<string> AddChartAsync(JsonObject? arguments, string path, int sheetIndex)
    {
        var chartTypeStr = arguments?["chartType"]?.GetValue<string>() ?? throw new ArgumentException("chartType is required for add operation");
        var dataRange = arguments?["dataRange"]?.GetValue<string>() ?? throw new ArgumentException("dataRange is required for add operation");
        var categoryAxisDataRange = arguments?["categoryAxisDataRange"]?.GetValue<string>();
        var title = arguments?["title"]?.GetValue<string>();
        var topRow = arguments?["topRow"]?.GetValue<int>();
        var leftColumn = arguments?["leftColumn"]?.GetValue<int>() ?? 0;
        var width = arguments?["width"]?.GetValue<int>() ?? 10;
        var height = arguments?["height"]?.GetValue<int>() ?? 15;

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

        int chartTopRow;
        if (topRow.HasValue)
        {
            chartTopRow = topRow.Value;
        }
        else
        {
            var dataRangeObj = worksheet.Cells.CreateRange(dataRange.Split(',')[0].Trim());
            chartTopRow = dataRangeObj.FirstRow + dataRangeObj.RowCount + 2;
        }

        int chartIndex = worksheet.Charts.Add(chartType, chartTopRow, leftColumn, chartTopRow + height, leftColumn + width);
        var chart = worksheet.Charts[chartIndex];
        
        // Clear existing series
        chart.NSeries.Clear();
        
        // Parse data range - support multiple ranges separated by comma (e.g., "E2:E40,G2:G40")
        var ranges = dataRange.Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
        
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
        workbook.Save(path);

        return await Task.FromResult($"Chart added to worksheet with data range: {dataRange}");
    }

    private async Task<string> EditChartAsync(JsonObject? arguments, string path, string outputPath, int sheetIndex)
    {
        var chartIndex = arguments?["chartIndex"]?.GetValue<int>() ?? throw new ArgumentException("chartIndex is required for edit operation");
        var title = arguments?["title"]?.GetValue<string>();
        var dataRange = arguments?["dataRange"]?.GetValue<string>();
        var categoryAxisDataRange = arguments?["categoryAxisDataRange"]?.GetValue<string>();
        var chartTypeStr = arguments?["chartType"]?.GetValue<string>();
        var showLegend = arguments?["showLegend"]?.GetValue<bool?>();
        var legendPosition = arguments?["legendPosition"]?.GetValue<string>();

        using var workbook = new Workbook(path);
        
        var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);
        var charts = worksheet.Charts;
        
        if (chartIndex < 0 || chartIndex >= charts.Count)
        {
            throw new ArgumentException($"圖表索引 {chartIndex} 超出範圍 (工作表共有 {charts.Count} 個圖表)");
        }

        var chart = charts[chartIndex];
        var changes = new List<string>();

        if (!string.IsNullOrEmpty(title))
        {
            chart.Title.Text = title;
            changes.Add($"標題: {title}");
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
                rangeInfo += $", X軸: {categoryAxisDataRange}";
            }
            changes.Add($"數據範圍: {rangeInfo}");
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
                _ => chart.Type
            };
            chart.Type = chartType;
            changes.Add($"圖表類型: {chartTypeStr}");
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
            changes.Add($"圖例: {(showLegend.Value ? "顯示" : "隱藏")}");
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

    private async Task<string> DeleteChartAsync(JsonObject? arguments, string path, int sheetIndex)
    {
        var chartIndex = arguments?["chartIndex"]?.GetValue<int>() ?? throw new ArgumentException("chartIndex is required for delete operation");

        using var workbook = new Workbook(path);
        
        var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);
        var charts = worksheet.Charts;
        
        if (chartIndex < 0 || chartIndex >= charts.Count)
        {
            throw new ArgumentException($"圖表索引 {chartIndex} 超出範圍 (工作表共有 {charts.Count} 個圖表)");
        }

        var chart = charts[chartIndex];
        var chartName = chart.Name ?? $"圖表 {chartIndex}";
        
        charts.RemoveAt(chartIndex);
        workbook.Save(path);
        
        var remainingCount = charts.Count;
        
        return await Task.FromResult($"成功刪除圖表 #{chartIndex} ({chartName})\n工作表剩餘圖表數: {remainingCount}\n輸出: {path}");
    }

    private async Task<string> GetChartsAsync(JsonObject? arguments, string path, int sheetIndex)
    {
        using var workbook = new Workbook(path);
        
        var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);
        var charts = worksheet.Charts;
        var result = new StringBuilder();

        result.AppendLine($"=== 工作表 '{worksheet.Name}' 的圖表資訊 ===\n");
        result.AppendLine($"總圖表數: {charts.Count}\n");

        if (charts.Count == 0)
        {
            result.AppendLine("未找到圖表");
            return await Task.FromResult(result.ToString());
        }

        for (int i = 0; i < charts.Count; i++)
        {
            var chart = charts[i];
            result.AppendLine($"【圖表 {i}】");
            result.AppendLine($"名稱: {chart.Name ?? "(無名稱)"}");
            result.AppendLine($"類型: {chart.Type}");
            result.AppendLine($"位置: 行 {chart.ChartObject.UpperLeftRow} - {chart.ChartObject.LowerRightRow}, 列 {chart.ChartObject.UpperLeftColumn} - {chart.ChartObject.LowerRightColumn}");
            result.AppendLine($"寬度: {chart.ChartObject.Width}");
            result.AppendLine($"高度: {chart.ChartObject.Height}");
            
            if (chart.NSeries != null && chart.NSeries.Count > 0)
            {
                result.AppendLine($"數據系列數: {chart.NSeries.Count}");
                for (int j = 0; j < chart.NSeries.Count && j < 5; j++)
                {
                    var series = chart.NSeries[j];
                    var seriesName = series.Name ?? "(無名稱)";
                    var valuesRange = series.Values ?? "";
                    
                    result.AppendLine($"  系列 {j}: {seriesName}");
                    if (!string.IsNullOrEmpty(valuesRange))
                    {
                        result.AppendLine($"    數據範圍 (Y軸): {valuesRange}");
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
                        result.AppendLine($"    X軸數據: {categoryData}");
                    }
                }
            }
            
            if (chart.Title != null)
            {
                result.AppendLine($"標題: {chart.Title.Text}");
            }
            
            if (chart.Legend != null)
            {
                result.AppendLine($"圖例: 已啟用");
            }
            result.AppendLine();
        }

        return await Task.FromResult(result.ToString());
    }

    private async Task<string> UpdateChartDataAsync(JsonObject? arguments, string path, string outputPath, int sheetIndex)
    {
        var chartIndex = arguments?["chartIndex"]?.GetValue<int>() ?? throw new ArgumentException("chartIndex is required for update_data operation");
        var dataRange = arguments?["dataRange"]?.GetValue<string>() ?? throw new ArgumentException("dataRange is required for update_data operation");

        using var workbook = new Workbook(path);
        
        var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);
        var charts = worksheet.Charts;
        
        if (chartIndex < 0 || chartIndex >= charts.Count)
        {
            throw new ArgumentException($"圖表索引 {chartIndex} 超出範圍 (工作表共有 {charts.Count} 個圖表)");
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

        return await Task.FromResult($"成功更新圖表 #{chartIndex} 的數據源\n新數據範圍: {dataRange}\n輸出: {outputPath}");
    }

    private async Task<string> SetChartPropertiesAsync(JsonObject? arguments, string path, string outputPath, int sheetIndex)
    {
        var chartIndex = arguments?["chartIndex"]?.GetValue<int>() ?? throw new ArgumentException("chartIndex is required for set_properties operation");
        var title = arguments?["title"]?.GetValue<string>();
        var removeTitle = arguments?["removeTitle"]?.GetValue<bool?>() ?? false;
        var legendVisible = arguments?["legendVisible"]?.GetValue<bool?>();
        var legendPosition = arguments?["legendPosition"]?.GetValue<string>();

        using var workbook = new Workbook(path);
        var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);
        PowerPointHelper.ValidateCollectionIndex(chartIndex, worksheet.Charts, "圖表");

        var chart = worksheet.Charts[chartIndex];
        var changes = new List<string>();

        // Handle title
        if (removeTitle)
        {
            chart.Title.Text = "";
            changes.Add("標題已移除");
        }
        else if (!string.IsNullOrEmpty(title))
        {
            chart.Title.Text = title;
            changes.Add($"標題: {title}");
        }

        // Handle legend visibility
        if (legendVisible.HasValue)
        {
            chart.ShowLegend = legendVisible.Value;
            changes.Add($"圖例: {(legendVisible.Value ? "顯示" : "隱藏")}");
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
            changes.Add($"圖例位置: {legendPosition}");
        }

        workbook.Save(outputPath);
        
        var result = changes.Count > 0 
            ? $"圖表屬性已更新: {string.Join(", ", changes)}\n輸出: {outputPath}"
            : $"無變更\n輸出: {outputPath}";
        
        return await Task.FromResult(result);
    }

}

