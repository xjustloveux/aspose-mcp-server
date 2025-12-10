using System.Text.Json.Nodes;
using Aspose.Words;
using Aspose.Words.Fields;

namespace AsposeMcpServer.Tools;

public class WordAddHyperlinkTool : IAsposeTool
{
    public string Description => "Add a hyperlink to a Word document";

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
            text = new
            {
                type = "string",
                description = "Display text for the hyperlink"
            },
            url = new
            {
                type = "string",
                description = "URL or target address for the hyperlink (e.g., 'https://example.com' or '#bookmark' for internal bookmark)"
            },
            paragraphIndex = new
            {
                type = "number",
                description = "Index of the paragraph to insert hyperlink after (0-based). If not provided, inserts at the end. Use -1 to insert at the beginning."
            },
            tooltip = new
            {
                type = "string",
                description = "Tooltip text shown when hovering over the hyperlink"
            }
        },
        required = new[] { "path", "text", "url" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var outputPath = arguments?["outputPath"]?.GetValue<string>() ?? path;
        var text = arguments?["text"]?.GetValue<string>() ?? throw new ArgumentException("text is required");
        var url = arguments?["url"]?.GetValue<string>() ?? throw new ArgumentException("url is required");
        var paragraphIndex = arguments?["paragraphIndex"]?.GetValue<int?>();
        var tooltip = arguments?["tooltip"]?.GetValue<string>();

        var doc = new Document(path);
        var builder = new DocumentBuilder(doc);
        
        // Determine insertion position
        if (paragraphIndex.HasValue)
        {
            var paragraphs = doc.GetChildNodes(NodeType.Paragraph, true);
            if (paragraphIndex.Value == -1)
            {
                // Insert at the beginning
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
                // Insert after the specified paragraph
                var targetPara = paragraphs[paragraphIndex.Value] as Paragraph;
                if (targetPara != null)
                {
                    builder.MoveTo(targetPara);
                }
                else
                {
                    throw new InvalidOperationException($"無法找到索引 {paragraphIndex.Value} 的段落");
                }
            }
            else
            {
                throw new ArgumentException($"段落索引 {paragraphIndex.Value} 超出範圍 (文檔共有 {paragraphs.Count} 個段落)");
            }
        }
        else
        {
            // Default: Move to end of document
            builder.MoveToDocumentEnd();
        }
        
        // Insert hyperlink
        builder.InsertHyperlink(text, url, false);
        
        // Set tooltip if provided
        if (!string.IsNullOrEmpty(tooltip))
        {
            // Get the last inserted field (should be the hyperlink field)
            var fields = doc.Range.Fields;
            if (fields.Count > 0)
            {
                var lastField = fields[fields.Count - 1];
                if (lastField is FieldHyperlink hyperlinkField)
                {
                    // Set screen tip (tooltip) - this is stored in the field result
                    // Note: Aspose.Words may handle this differently, but we can try setting it
                    hyperlinkField.ScreenTip = tooltip;
                }
            }
        }
        
        doc.Save(outputPath);
        
        var result = $"成功添加超連結\n";
        result += $"顯示文字: {text}\n";
        result += $"URL: {url}\n";
        if (!string.IsNullOrEmpty(tooltip))
        {
            result += $"提示文字: {tooltip}\n";
        }
        if (paragraphIndex.HasValue)
        {
            if (paragraphIndex.Value == -1)
            {
                result += "插入位置: 文檔開頭\n";
            }
            else
            {
                result += $"插入位置: 段落 #{paragraphIndex.Value} 之後\n";
            }
        }
        else
        {
            result += "插入位置: 文檔末尾\n";
        }
        result += $"輸出: {outputPath}";
        
        return await Task.FromResult(result);
    }
}

