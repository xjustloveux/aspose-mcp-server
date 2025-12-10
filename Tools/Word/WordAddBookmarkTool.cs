using System.Text.Json.Nodes;
using Aspose.Words;

namespace AsposeMcpServer.Tools;

public class WordAddBookmarkTool : IAsposeTool
{
    public string Description => "Add a bookmark to a Word document";

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
            name = new
            {
                type = "string",
                description = "Bookmark name (must be unique)"
            },
            text = new
            {
                type = "string",
                description = "Text content to bookmark (optional, if not provided, bookmark will be inserted at document end)"
            },
            paragraphIndex = new
            {
                type = "number",
                description = "Paragraph index to insert bookmark at (0-based). If not provided, inserts at the end. Use -1 to insert at the beginning."
            }
        },
        required = new[] { "path", "name" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var outputPath = arguments?["outputPath"]?.GetValue<string>() ?? path;
        var name = arguments?["name"]?.GetValue<string>() ?? throw new ArgumentException("name is required");
        var text = arguments?["text"]?.GetValue<string>();
        var paragraphIndex = arguments?["paragraphIndex"]?.GetValue<int?>();

        var doc = new Document(path);
        var builder = new DocumentBuilder(doc);
        
        // Check if bookmark already exists
        if (doc.Range.Bookmarks[name] != null)
        {
            throw new InvalidOperationException($"書籤 '{name}' 已存在");
        }
        
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
        
        // Insert bookmark
        builder.StartBookmark(name);
        
        // Add text if provided
        if (!string.IsNullOrEmpty(text))
        {
            builder.Write(text);
        }
        
        builder.EndBookmark(name);
        
        doc.Save(outputPath);
        
        var result = $"成功添加書籤\n";
        result += $"書籤名稱: {name}\n";
        if (!string.IsNullOrEmpty(text))
        {
            result += $"書籤文字: {text}\n";
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

