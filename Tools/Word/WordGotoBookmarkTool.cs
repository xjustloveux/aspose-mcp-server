using System.Text.Json.Nodes;
using Aspose.Words;

namespace AsposeMcpServer.Tools;

public class WordGotoBookmarkTool : IAsposeTool
{
    public string Description => "Get bookmark location information (goto bookmark)";

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
            name = new
            {
                type = "string",
                description = "Bookmark name"
            }
        },
        required = new[] { "path", "name" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var name = arguments?["name"]?.GetValue<string>() ?? throw new ArgumentException("name is required");

        var doc = new Document(path);
        
        // Get bookmark
        var bookmark = doc.Range.Bookmarks[name];
        if (bookmark == null)
        {
            throw new ArgumentException($"找不到書籤 '{name}'，可用書籤請使用 word_get_bookmarks 工具查看");
        }
        
        // Get bookmark information
        string bookmarkText = bookmark.Text;
        var bookmarkRange = bookmark.BookmarkStart?.ParentNode as Paragraph;
        
        // Try to find paragraph index
        int paragraphIndex = -1;
        var paragraphs = doc.GetChildNodes(NodeType.Paragraph, true);
        for (int i = 0; i < paragraphs.Count; i++)
        {
            if (paragraphs[i] == bookmarkRange)
            {
                paragraphIndex = i;
                break;
            }
        }
        
        var result = $"書籤位置資訊\n";
        result += $"書籤名稱: {name}\n";
        result += $"書籤文字: {bookmarkText}\n";
        if (paragraphIndex >= 0)
        {
            result += $"段落索引: {paragraphIndex}\n";
        }
        result += $"書籤範圍長度: {bookmarkText.Length} 個字元\n";
        
        // Get surrounding context if possible
        if (bookmarkRange != null)
        {
            string paraText = bookmarkRange.GetText().Trim();
            if (paraText.Length > 100)
            {
                paraText = paraText.Substring(0, 100) + "...";
            }
            result += $"所在段落內容: {paraText}\n";
        }
        
        return await Task.FromResult(result);
    }
}

