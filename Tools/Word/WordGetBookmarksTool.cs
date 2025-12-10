using System.Text.Json.Nodes;
using Aspose.Words;

namespace AsposeMcpServer.Tools;

public class WordGetBookmarksTool : IAsposeTool
{
    public string Description => "Get all bookmarks from Word document";

    public object InputSchema => new
    {
        type = "object",
        properties = new
        {
            path = new
            {
                type = "string",
                description = "Input file path"
            }
        },
        required = new[] { "path" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");

        var doc = new Document(path);
        
        // Get all bookmarks
        var bookmarks = doc.Range.Bookmarks;
        
        if (bookmarks.Count == 0)
        {
            return await Task.FromResult("文檔中沒有找到書籤");
        }
        
        var result = new System.Text.StringBuilder();
        result.AppendLine($"找到 {bookmarks.Count} 個書籤：\n");
        
        int index = 0;
        foreach (Bookmark bookmark in bookmarks)
        {
            result.AppendLine($"書籤 #{index}:");
            result.AppendLine($"  名稱: {bookmark.Name}");
            result.AppendLine($"  文字: {bookmark.Text}");
            result.AppendLine($"  長度: {bookmark.Text.Length} 個字元");
            result.AppendLine();
            index++;
        }
        
        return await Task.FromResult(result.ToString().TrimEnd());
    }
}

