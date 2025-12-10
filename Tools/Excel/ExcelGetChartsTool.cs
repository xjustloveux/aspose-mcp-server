using System.Text.Json.Nodes;
using System.Text;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeMcpServer.Tools;

public class ExcelGetChartsTool : IAsposeTool
{
    public string Description => "Get all charts information from an Excel worksheet";

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
            }
        },
        required = new[] { "path" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var sheetIndex = arguments?["sheetIndex"]?.GetValue<int>() ?? 0;

        using var workbook = new Workbook(path);
        
        if (sheetIndex < 0 || sheetIndex >= workbook.Worksheets.Count)
        {
            throw new ArgumentException($"工作表索引 {sheetIndex} 超出範圍 (共有 {workbook.Worksheets.Count} 個工作表)");
        }

        var worksheet = workbook.Worksheets[sheetIndex];
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
                for (int j = 0; j < chart.NSeries.Count && j < 5; j++) // Limit to first 5 series
                {
                    var series = chart.NSeries[j];
                    var seriesName = series.Name ?? "(無名稱)";
                    var valuesRange = series.Values ?? "";
                    
                    result.AppendLine($"  系列 {j}: {seriesName}");
                    if (!string.IsNullOrEmpty(valuesRange))
                    {
                        result.AppendLine($"    數據範圍 (Y軸): {valuesRange}");
                    }
                    
                    // Try to get category axis (X-axis) data using multiple methods
                    // Priority: Try to get full range from chart data source first (Method 3),
                    // then try other methods if that fails
                    string? categoryData = null;
                    
                    // Method 3 (Priority): Try to get from chart's data source first
                    // This should give us the full range (e.g., A2:A40) rather than single cell
                    try
                    {
                        // Check if chart has a data source that includes category data
                        var chartDataRange = chart.GetChartDataRange();
                        if (!string.IsNullOrEmpty(chartDataRange))
                        {
                            // If chart data range contains comma, first part might be category
                            var parts = chartDataRange.Split(',');
                            if (parts.Length >= 2)
                            {
                                // First part is typically category axis (full range)
                                var potentialCategory = parts[0].Trim();
                                // Prefer full range (contains :) over single cell reference
                                if (potentialCategory.Contains(":") || potentialCategory.Contains("$"))
                                {
                                    categoryData = potentialCategory;
                                }
                            }
                        }
                    }
                    catch { }
                    
                    // Method 1: Try CategoryData property (direct property access)
                    // Only use if Method 3 didn't find a full range
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
                                    // Prefer full range over single cell
                                    if (!string.IsNullOrEmpty(catDataStr))
                                    {
                                        if (catDataStr.Contains(":"))
                                        {
                                            // Full range - use it
                                            categoryData = catDataStr;
                                        }
                                        else if (string.IsNullOrEmpty(categoryData) && 
                                                 (catDataStr.Contains("$") || catDataStr.StartsWith("=")))
                                        {
                                            // Single cell - use only if no other data found
                                            categoryData = catDataStr;
                                        }
                                    }
                                }
                            }
                        }
                        catch { }
                    }
                    
                    // Method 2: Try XValues property (for some chart types like scatter)
                    // Only use if we don't have a full range yet
                    if (string.IsNullOrEmpty(categoryData) || !categoryData.Contains(":"))
                    {
                        try
                        {
                            var xValuesProp = series.GetType().GetProperty("XValues");
                            if (xValuesProp != null)
                            {
                                var xValues = xValuesProp.GetValue(series);
                                if (xValues != null)
                                {
                                    var xValuesStr = xValues.ToString();
                                    if (!string.IsNullOrEmpty(xValuesStr))
                                    {
                                        if (xValuesStr.Contains(":"))
                                        {
                                            // Full range - use it
                                            categoryData = xValuesStr;
                                        }
                                        else if (string.IsNullOrEmpty(categoryData) && 
                                                 (xValuesStr.Contains("$") || xValuesStr.StartsWith("=")))
                                        {
                                            // Single cell - use only if no other data found
                                            categoryData = xValuesStr;
                                        }
                                    }
                                }
                            }
                        }
                        catch { }
                    }
                    
                    // Method 4: Try CategoryAxisData property
                    if (string.IsNullOrEmpty(categoryData))
                    {
                        try
                        {
                            var catAxisDataProp = series.GetType().GetProperty("CategoryAxisData");
                            if (catAxisDataProp != null)
                            {
                                var catAxisData = catAxisDataProp.GetValue(series);
                                if (catAxisData != null)
                                {
                                    var catAxisDataStr = catAxisData.ToString();
                                    if (!string.IsNullOrEmpty(catAxisDataStr) && 
                                        (catAxisDataStr.Contains("$") || catAxisDataStr.Contains(":") || catAxisDataStr.StartsWith("=")))
                                    {
                                        categoryData = catAxisDataStr;
                                    }
                                }
                            }
                        }
                        catch { }
                    }
                    
                    // Method 5: Try to inspect all properties of Series to find category-related ones
                    if (string.IsNullOrEmpty(categoryData))
                    {
                        try
                        {
                            var props = series.GetType().GetProperties();
                            foreach (var prop in props)
                            {
                                if (prop.Name.Contains("Category") || prop.Name.Contains("XAxis") || prop.Name.Contains("XValue"))
                                {
                                    var value = prop.GetValue(series);
                                    if (value != null)
                                    {
                                        var valueStr = value.ToString();
                                        if (!string.IsNullOrEmpty(valueStr) && 
                                            (valueStr.Contains("$") || valueStr.Contains(":") || valueStr.StartsWith("=")) &&
                                            !valueStr.StartsWith("System."))
                                        {
                                            categoryData = valueStr;
                                            break;
                                        }
                                    }
                                }
                            }
                        }
                        catch { }
                    }
                    
                    // Display X-axis data if found
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
}

