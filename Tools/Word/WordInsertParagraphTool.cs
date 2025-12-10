using System.Text.Json.Nodes;
using Aspose.Words;

namespace AsposeMcpServer.Tools;

public class WordInsertParagraphTool : IAsposeTool
{
    public string Description => "Insert a new paragraph at a specific position in Word document";

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
            paragraphIndex = new
            {
                type = "number",
                description = "Index of the paragraph to insert after (0-based). Use -1 to insert at the beginning. If not provided, inserts at the end."
            },
            text = new
            {
                type = "string",
                description = "Text content for the paragraph"
            },
            styleName = new
            {
                type = "string",
                description = "Style name to apply (e.g., 'Heading 1', '標題1', 'Normal')"
            },
            alignment = new
            {
                type = "string",
                description = "Text alignment: left, center, right, justify",
                @enum = new[] { "left", "center", "right", "justify" }
            }
        },
        required = new[] { "path", "text" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var outputPath = arguments?["outputPath"]?.GetValue<string>() ?? path;
        var text = arguments?["text"]?.GetValue<string>() ?? throw new ArgumentException("text is required");
        var paragraphIndex = arguments?["paragraphIndex"]?.GetValue<int?>();
        var styleName = arguments?["styleName"]?.GetValue<string>();
        var alignment = arguments?["alignment"]?.GetValue<string>();

        var doc = new Document(path);
        var paragraphs = doc.GetChildNodes(NodeType.Paragraph, true);
        
        Paragraph? targetPara = null;
        string insertPosition = "文檔末尾";
        
        // Determine insertion position
        if (paragraphIndex.HasValue)
        {
            if (paragraphIndex.Value == -1)
            {
                // Insert at the beginning
                if (paragraphs.Count > 0)
                {
                    targetPara = paragraphs[0] as Paragraph;
                    insertPosition = "文檔開頭";
                }
            }
            else if (paragraphIndex.Value >= 0 && paragraphIndex.Value < paragraphs.Count)
            {
                // Insert after the specified paragraph
                targetPara = paragraphs[paragraphIndex.Value] as Paragraph;
                insertPosition = $"段落 #{paragraphIndex.Value} 之後";
            }
            else
            {
                throw new ArgumentException($"段落索引 {paragraphIndex.Value} 超出範圍 (文檔共有 {paragraphs.Count} 個段落)");
            }
        }

        // Create new paragraph
        var para = new Paragraph(doc);
        var run = new Run(doc, text);
        para.AppendChild(run);

        // Apply style if provided
        if (!string.IsNullOrEmpty(styleName))
        {
            try
            {
                var style = doc.Styles[styleName];
                if (style != null)
                {
                    para.ParagraphFormat.StyleName = styleName;
                }
                else
                {
                    throw new ArgumentException($"找不到樣式 '{styleName}'，可用樣式請使用 word_get_styles 工具查看");
                }
            }
            catch (ArgumentException)
            {
                throw;
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException($"無法應用樣式 '{styleName}': {ex.Message}，可用樣式請使用 word_get_styles 工具查看", ex);
            }
        }

        // Apply alignment if provided
        if (!string.IsNullOrEmpty(alignment))
        {
            para.ParagraphFormat.Alignment = alignment.ToLower() switch
            {
                "left" => ParagraphAlignment.Left,
                "right" => ParagraphAlignment.Right,
                "center" => ParagraphAlignment.Center,
                "justify" => ParagraphAlignment.Justify,
                _ => ParagraphAlignment.Left
            };
        }

        // Insert paragraph at the correct position
        if (targetPara != null)
        {
            if (paragraphIndex!.Value == -1)
            {
                // Insert at the beginning - insert before first paragraph
                targetPara.ParentNode.InsertBefore(para, targetPara);
            }
            else
            {
                // Insert after the specified paragraph
                targetPara.ParentNode.InsertAfter(para, targetPara);
            }
        }
        else
        {
            // Default: Append to end
            var body = doc.FirstSection.Body;
            body.AppendChild(para);
        }

        doc.Save(outputPath);

        var result = $"成功插入段落\n";
        result += $"插入位置: {insertPosition}\n";
        if (!string.IsNullOrEmpty(styleName))
        {
            result += $"應用樣式: {styleName}\n";
        }
        if (!string.IsNullOrEmpty(alignment))
        {
            result += $"對齊方式: {alignment}\n";
        }
        result += $"文檔段落數: {doc.GetChildNodes(NodeType.Paragraph, true).Count}\n";
        result += $"輸出: {outputPath}";

        return await Task.FromResult(result);
    }
}

