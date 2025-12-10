using System.Text.Json.Nodes;
using Aspose.Words;

namespace AsposeMcpServer.Tools;

public class WordDeleteBookmarkTool : IAsposeTool
{
    public string Description => "Delete a bookmark from Word document";

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
                description = "Bookmark name to delete"
            },
            keepText = new
            {
                type = "boolean",
                description = "Keep bookmark text content (default: true). If false, removes both bookmark and its text."
            }
        },
        required = new[] { "path", "name" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var outputPath = arguments?["outputPath"]?.GetValue<string>() ?? path;
        var name = arguments?["name"]?.GetValue<string>() ?? throw new ArgumentException("name is required");
        var keepText = arguments?["keepText"]?.GetValue<bool>() ?? true;

        var doc = new Document(path);
        
        // Get bookmark
        var bookmark = doc.Range.Bookmarks[name];
        if (bookmark == null)
        {
            throw new ArgumentException($"找不到書籤 '{name}'，可用書籤請使用 word_get_bookmarks 工具查看");
        }
        
        // Get bookmark info before deletion
        string bookmarkText = bookmark.Text;
        
        // Delete bookmark
        if (keepText)
        {
            // Remove bookmark markers but keep text
            bookmark.Remove();
        }
        else
        {
            // Remove bookmark and its text content
            bookmark.BookmarkStart?.Remove();
            bookmark.BookmarkEnd?.Remove();
        }
        
        doc.Save(outputPath);
        
        // Count remaining bookmarks
        int remainingCount = doc.Range.Bookmarks.Count;
        
        var result = $"成功刪除書籤 '{name}'\n";
        result += $"書籤文字: {bookmarkText}\n";
        result += $"保留文字: {(keepText ? "是" : "否")}\n";
        result += $"文檔剩餘書籤數: {remainingCount}\n";
        result += $"輸出: {outputPath}";
        
        return await Task.FromResult(result);
    }
}

