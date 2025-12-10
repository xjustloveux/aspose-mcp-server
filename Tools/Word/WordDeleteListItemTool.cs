using System.Text.Json.Nodes;
using Aspose.Words;

namespace AsposeMcpServer.Tools;

public class WordDeleteListItemTool : IAsposeTool
{
    public string Description => "Delete a specific list item from Word document by paragraph index";

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
                description = "Paragraph index of the list item to delete (0-based)"
            }
        },
        required = new[] { "path", "paragraphIndex" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var outputPath = arguments?["outputPath"]?.GetValue<string>() ?? path;
        var paragraphIndex = arguments?["paragraphIndex"]?.GetValue<int>() ?? throw new ArgumentException("paragraphIndex is required");

        var doc = new Document(path);
        var paragraphs = doc.GetChildNodes(NodeType.Paragraph, true);
        
        if (paragraphIndex < 0 || paragraphIndex >= paragraphs.Count)
        {
            throw new ArgumentException($"段落索引 {paragraphIndex} 超出範圍 (文檔共有 {paragraphs.Count} 個段落)");
        }
        
        var paraToDelete = paragraphs[paragraphIndex] as Paragraph;
        if (paraToDelete == null)
        {
            throw new InvalidOperationException($"無法獲取索引 {paragraphIndex} 的段落");
        }
        
        // Get list item info before deletion
        string itemText = paraToDelete.GetText().Trim();
        string itemPreview = "";
        if (itemText.Length > 50)
        {
            itemPreview = itemText.Substring(0, 50) + "...";
        }
        else
        {
            itemPreview = itemText;
        }
        
        // Check if it's a list item (has list formatting)
        bool isListItem = paraToDelete.ListFormat.IsListItem;
        string listInfo = isListItem ? "（清單項目）" : "（一般段落）";
        
        // Delete the paragraph
        paraToDelete.Remove();
        
        doc.Save(outputPath);
        
        var result = $"成功刪除清單項目 #{paragraphIndex}{listInfo}\n";
        if (!string.IsNullOrEmpty(itemPreview))
        {
            result += $"內容預覽: {itemPreview}\n";
        }
        result += $"文檔剩餘段落數: {doc.GetChildNodes(NodeType.Paragraph, true).Count}\n";
        result += $"輸出: {outputPath}";
        
        return await Task.FromResult(result);
    }
}

