using System.Text.Json.Nodes;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;
using System.IO;

namespace AsposeMcpServer.Tools;

public class WordAddChartTool : IAsposeTool
{
    public string Description => "Add a chart (bar chart, line chart, pie chart, etc.) to Word document by embedding Excel chart";

    public object InputSchema => new
    {
        type = "object",
        properties = new
        {
            path = new
            {
                type = "string",
                description = "Input file path"
            },
            outputPath = new
            {
                type = "string",
                description = "Output file path (if not provided, overwrites input)"
            },
            chartType = new
            {
                type = "string",
                description = "Chart type: column, bar, line, pie, area, scatter, doughnut (default: column)",
                @enum = new[] { "column", "bar", "line", "pie", "area", "scatter", "doughnut" }
            },
            data = new
            {
                type = "array",
                description = "Chart data as 2D array. First row can be headers, subsequent rows are data values. Example: [[\"Category\", \"Value\"], [\"A\", 10], [\"B\", 20], [\"C\", 30]]",
                items = new
                {
                    type = "array",
                    items = new { type = "string" }
                }
            },
            title = new
            {
                type = "string",
                description = "Chart title (optional)"
            },
            width = new
            {
                type = "number",
                description = "Chart width in points (default: 432, which is 6 inches)"
            },
            height = new
            {
                type = "number",
                description = "Chart height in points (default: 252, which is 3.5 inches)"
            },
            paragraphIndex = new
            {
                type = "number",
                description = "Paragraph index to insert after (0-based). If not provided, inserts at end of document. Use -1 to insert at beginning."
            },
            alignment = new
            {
                type = "string",
                description = "Chart alignment: left, center, right (default: left)",
                @enum = new[] { "left", "center", "right" }
            }
        },
        required = new[] { "path", "data" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var outputPath = arguments?["outputPath"]?.GetValue<string>() ?? path;
        var chartType = arguments?["chartType"]?.GetValue<string>() ?? "column";
        var data = arguments?["data"]?.AsArray() ?? throw new ArgumentException("data is required");
        var title = arguments?["title"]?.GetValue<string>();
        var width = arguments?["width"]?.GetValue<double>() ?? 432.0; // 6 inches
        var height = arguments?["height"]?.GetValue<double>() ?? 252.0; // 3.5 inches
        var paragraphIndex = arguments?["paragraphIndex"]?.GetValue<int?>();
        var alignment = arguments?["alignment"]?.GetValue<string>() ?? "left";

        if (data.Count == 0)
        {
            throw new ArgumentException("圖表數據不能為空");
        }

        // Parse data
        var tableData = new List<List<string>>();
        foreach (var row in data)
        {
            if (row is JsonArray rowArray)
            {
                var rowData = new List<string>();
                foreach (var cell in rowArray)
                {
                    rowData.Add(cell?.ToString() ?? "");
                }
                tableData.Add(rowData);
            }
        }

        if (tableData.Count == 0)
        {
            throw new ArgumentException("無法解析圖表數據");
        }

        // Create temporary Excel file with chart
        string tempExcelPath = Path.Combine(Path.GetTempPath(), $"chart_{Guid.NewGuid()}.xlsx");
        try
        {
            var workbook = new Workbook();
            var worksheet = workbook.Worksheets[0];
            
            // Write data to Excel
            for (int i = 0; i < tableData.Count; i++)
            {
                for (int j = 0; j < tableData[i].Count; j++)
                {
                    var cellValue = tableData[i][j];
                    // Try to parse as number, otherwise use as string
                    if (double.TryParse(cellValue, out double numValue) && i > 0) // Skip header row
                    {
                        worksheet.Cells[i, j].PutValue(numValue);
                    }
                    else
                    {
                        worksheet.Cells[i, j].PutValue(cellValue);
                    }
                }
            }
            
            // Determine data range
            int maxCol = tableData.Max(r => r.Count);
            string dataRange = $"A1:{Convert.ToChar(64 + maxCol)}{tableData.Count}";
            
            // Create chart
            var chartTypeEnum = chartType.ToLower() switch
            {
                "bar" => ChartType.Bar,
                "line" => ChartType.Line,
                "pie" => ChartType.Pie,
                "area" => ChartType.Area,
                "scatter" => ChartType.Scatter,
                "doughnut" => ChartType.Doughnut,
                _ => ChartType.Column
            };
            
            int chartIndex = worksheet.Charts.Add(chartTypeEnum, 0, tableData.Count + 2, 20, 10);
            var chart = worksheet.Charts[chartIndex];
            chart.SetChartDataRange(dataRange, true);
            
            if (!string.IsNullOrEmpty(title))
            {
                chart.Title.Text = title;
            }
            
            // Save temporary Excel file
            workbook.Save(tempExcelPath);
            
            // Convert chart to image
            string tempImagePath = Path.Combine(Path.GetTempPath(), $"chart_{Guid.NewGuid()}.png");
            
            // Render chart to image
            chart.ToImage(tempImagePath, Aspose.Cells.Drawing.ImageType.Png);
            
            // Now insert the image into Word document
            var doc = new Document(path);
            var builder = new DocumentBuilder(doc);
            
            // Determine insertion position
            if (paragraphIndex.HasValue)
            {
                var paragraphs = doc.GetChildNodes(NodeType.Paragraph, true);
                if (paragraphIndex.Value == -1)
                {
                    if (paragraphs.Count > 0)
                    {
                        var firstPara = paragraphs[0] as Paragraph;
                        if (firstPara != null)
                        {
                            builder.MoveTo(firstPara);
                        }
                    }
                }
                else if (paragraphIndex.Value >= 0 && paragraphIndex.Value < paragraphs.Count)
                {
                    var targetPara = paragraphs[paragraphIndex.Value] as Paragraph;
                    if (targetPara != null)
                    {
                        builder.MoveTo(targetPara);
                    }
                    else
                    {
                        throw new ArgumentException($"無法找到索引 {paragraphIndex.Value} 的段落");
                    }
                }
                else
                {
                    throw new ArgumentException($"段落索引 {paragraphIndex.Value} 超出範圍 (文檔共有 {paragraphs.Count} 個段落)");
                }
            }
            else
            {
                builder.MoveToDocumentEnd();
            }
            
            // Set alignment
            builder.ParagraphFormat.Alignment = GetAlignment(alignment);
            
            // Insert chart image
            var shape = builder.InsertImage(tempImagePath);
            shape.Width = width;
            shape.Height = height;
            shape.WrapType = WrapType.Inline;
            
            // Clean up temporary image file
            if (File.Exists(tempImagePath))
            {
                try
                {
                    File.Delete(tempImagePath);
                }
                catch
                {
                    // Ignore cleanup errors
                }
            }
            
            doc.Save(outputPath);
            
            var result = $"成功添加圖表\n";
            result += $"圖表類型: {chartType}\n";
            if (!string.IsNullOrEmpty(title))
            {
                result += $"標題: {title}\n";
            }
            result += $"數據行數: {tableData.Count}\n";
            result += $"數據列數: {(tableData.Count > 0 ? tableData[0].Count : 0)}\n";
            result += $"輸出: {outputPath}";
            
            return await Task.FromResult(result);
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException($"創建圖表時發生錯誤: {ex.Message}", ex);
        }
        finally
        {
            // Clean up temporary file
            if (File.Exists(tempExcelPath))
            {
                try
                {
                    File.Delete(tempExcelPath);
                }
                catch
                {
                    // Ignore cleanup errors
                }
            }
        }
    }
    
    private ParagraphAlignment GetAlignment(string alignment)
    {
        return alignment.ToLower() switch
        {
            "center" => ParagraphAlignment.Center,
            "right" => ParagraphAlignment.Right,
            _ => ParagraphAlignment.Left
        };
    }
}

