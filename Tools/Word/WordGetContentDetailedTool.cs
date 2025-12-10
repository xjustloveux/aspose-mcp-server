using System.Text;
using System.Text.Json;
using System.Text.Json.Nodes;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Tables;

namespace AsposeMcpServer.Tools;

public class WordGetContentDetailedTool : IAsposeTool
{
    public string Description => "Read detailed content from a Word document including headers, footers, styles, tables, and images";

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
            includeHeaders = new
            {
                type = "boolean",
                description = "Include headers (default: true)"
            },
            includeFooters = new
            {
                type = "boolean",
                description = "Include footers (default: true)"
            },
            includeStyles = new
            {
                type = "boolean",
                description = "Include style information (default: true)"
            },
            includeTables = new
            {
                type = "boolean",
                description = "Include table structure details (default: true)"
            },
            includeImages = new
            {
                type = "boolean",
                description = "Include image information (default: true)"
            }
        },
        required = new[] { "path" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var includeHeaders = arguments?["includeHeaders"]?.GetValue<bool>() ?? true;
        var includeFooters = arguments?["includeFooters"]?.GetValue<bool>() ?? true;
        var includeStyles = arguments?["includeStyles"]?.GetValue<bool>() ?? true;
        var includeTables = arguments?["includeTables"]?.GetValue<bool>() ?? true;
        var includeImages = arguments?["includeImages"]?.GetValue<bool>() ?? true;

        var doc = new Document(path);
        var result = new StringBuilder();

        // Document basic info
        result.AppendLine("=== 文檔基本信息 ===");
        result.AppendLine($"頁數: {doc.PageCount}");
        result.AppendLine($"段落數: {doc.GetChildNodes(NodeType.Paragraph, true).Count}");
        result.AppendLine($"表格數: {doc.GetChildNodes(NodeType.Table, true).Count}");
        result.AppendLine($"圖片數: {doc.GetChildNodes(NodeType.Shape, true).Cast<Shape>().Count(s => s.HasImage)}");
        result.AppendLine();

        // Headers
        if (includeHeaders)
        {
            result.AppendLine("=== 頁首 ===");
            foreach (Section section in doc.Sections)
            {
                var headerFooter = section.HeadersFooters[HeaderFooterType.HeaderPrimary];
                if (headerFooter != null && !string.IsNullOrWhiteSpace(headerFooter.ToString(SaveFormat.Text)))
                {
                    result.AppendLine($"Section {section.ParentNode.IndexOf(section) + 1} 頁首:");
                    result.AppendLine(headerFooter.ToString(SaveFormat.Text).Trim());
                    result.AppendLine();
                }
            }
        }

        // Footers
        if (includeFooters)
        {
            result.AppendLine("=== 頁尾 ===");
            foreach (Section section in doc.Sections)
            {
                var headerFooter = section.HeadersFooters[HeaderFooterType.FooterPrimary];
                if (headerFooter != null && !string.IsNullOrWhiteSpace(headerFooter.ToString(SaveFormat.Text)))
                {
                    result.AppendLine($"Section {section.ParentNode.IndexOf(section) + 1} 頁尾:");
                    result.AppendLine(headerFooter.ToString(SaveFormat.Text).Trim());
                    result.AppendLine();
                }
            }
        }

        // Main content with styles
        result.AppendLine("=== 正文內容 ===");
        foreach (Section section in doc.Sections)
        {
            foreach (Node node in section.Body)
            {
                if (node is Paragraph para)
                {
                    var text = para.ToString(SaveFormat.Text).Trim();
                    if (!string.IsNullOrWhiteSpace(text))
                    {
                        if (includeStyles)
                        {
                            result.Append($"[樣式: {para.ParagraphFormat.StyleName}");
                            if (para.Runs.Count > 0)
                            {
                                var run = para.Runs[0];
                                if (run.Font.Bold) result.Append(" 粗體");
                                if (run.Font.Italic) result.Append(" 斜體");
                                if (run.Font.Underline != Underline.None) result.Append(" 底線");
                                result.Append($" 字號:{run.Font.Size}");
                            }
                            result.Append("] ");
                        }
                        result.AppendLine(text);
                    }
                }
                else if (node is Table table && includeTables)
                {
                    result.AppendLine("\n[表格]");
                    result.AppendLine($"  行數: {table.Rows.Count}");
                    if (table.Rows.Count > 0)
                    {
                        result.AppendLine($"  列數: {table.Rows[0].Cells.Count}");
                    }
                    
                    // Table content
                    foreach (Row row in table.Rows)
                    {
                        result.Append("  | ");
                        foreach (Cell cell in row.Cells)
                        {
                            var cellText = cell.ToString(SaveFormat.Text).Trim().Replace("\r", "").Replace("\n", " ");
                            
                            if (includeStyles)
                            {
                                var shading = cell.CellFormat.Shading;
                                if (shading.BackgroundPatternColor.ToArgb() != System.Drawing.Color.Empty.ToArgb() &&
                                    shading.BackgroundPatternColor.ToArgb() != System.Drawing.Color.White.ToArgb())
                                {
                                    result.Append($"[背景色: {shading.BackgroundPatternColor.Name}] ");
                                }
                            }
                            
                            result.Append($"{cellText} | ");
                        }
                        result.AppendLine();
                    }
                    result.AppendLine();
                }
            }
        }

        // Images
        if (includeImages)
        {
            result.AppendLine("\n=== 圖片信息 ===");
            var shapes = doc.GetChildNodes(NodeType.Shape, true).Cast<Shape>().Where(s => s.HasImage).ToList();
            
            if (shapes.Any())
            {
                for (int i = 0; i < shapes.Count; i++)
                {
                    var shape = shapes[i];
                    result.AppendLine($"圖片 {i + 1}:");
                    result.AppendLine($"  寬度: {shape.Width} pt");
                    result.AppendLine($"  高度: {shape.Height} pt");
                    result.AppendLine($"  類型: {shape.ImageData.ImageType}");
                    
                    if (shape.Title != null && !string.IsNullOrWhiteSpace(shape.Title))
                    {
                        result.AppendLine($"  標題: {shape.Title}");
                    }
                    if (shape.AlternativeText != null && !string.IsNullOrWhiteSpace(shape.AlternativeText))
                    {
                        result.AppendLine($"  替代文字: {shape.AlternativeText}");
                    }
                    result.AppendLine();
                }
            }
            else
            {
                result.AppendLine("（無圖片）");
            }
        }

        return await Task.FromResult(result.ToString());
    }
}
