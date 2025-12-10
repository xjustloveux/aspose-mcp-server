using System.Text.Json.Nodes;
using Aspose.Words;

namespace AsposeMcpServer.Tools;

public class WordEditBookmarkTool : IAsposeTool
{
    public string Description => "Edit bookmark name or content in a Word document";

    public object InputSchema => new
    {
        type = "object",
        properties = new
        {
            path = new
            {
                type = "string",
                description = "Document file path"
            },
            outputPath = new
            {
                type = "string",
                description = "Output file path (if not provided, overwrites input)"
            },
            bookmarkName = new
            {
                type = "string",
                description = "Bookmark name to edit"
            },
            newName = new
            {
                type = "string",
                description = "New bookmark name (optional)"
            },
            newText = new
            {
                type = "string",
                description = "New text content for the bookmark (optional, replaces existing content)"
            }
        },
        required = new[] { "path", "bookmarkName" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var outputPath = arguments?["outputPath"]?.GetValue<string>() ?? path;
        var bookmarkName = arguments?["bookmarkName"]?.GetValue<string>() ?? throw new ArgumentException("bookmarkName is required");
        var newName = arguments?["newName"]?.GetValue<string>();
        var newText = arguments?["newText"]?.GetValue<string>();

        var doc = new Document(path);
        var bookmarks = doc.Range.Bookmarks;
        
        Bookmark? bookmark = null;
        try
        {
            bookmark = bookmarks[bookmarkName];
        }
        catch
        {
            throw new ArgumentException($"找不到書籤 '{bookmarkName}'");
        }
        
        if (bookmark == null)
        {
            throw new ArgumentException($"找不到書籤 '{bookmarkName}'");
        }
        var oldName = bookmark.Name;
        var oldText = bookmark.Text;
        var changes = new List<string>();
        
        // Rename bookmark if new name provided
        if (!string.IsNullOrEmpty(newName) && newName != bookmarkName)
        {
            // Check if new name already exists
            Bookmark? existingBookmark = null;
            try
            {
                existingBookmark = bookmarks[newName];
            }
            catch
            {
                // New name doesn't exist, continue
            }
            
            if (existingBookmark != null)
            {
                throw new ArgumentException($"書籤名稱 '{newName}' 已存在");
            }
            
            try
            {
                bookmark.Name = newName;
                changes.Add($"書籤名稱: {oldName} -> {newName}");
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException($"無法重命名書籤: {ex.Message}", ex);
            }
        }
        
        // Update bookmark text if new text provided
        if (!string.IsNullOrEmpty(newText))
        {
            try
            {
                // Get the bookmark range and replace its content
                var bookmarkRange = bookmark.BookmarkStart;
                var bookmarkEnd = bookmark.BookmarkEnd;
                
                if (bookmarkRange != null && bookmarkEnd != null)
                {
                    // Remove existing content between bookmark start and end
                    var currentNode = bookmarkRange.NextSibling;
                    while (currentNode != null && currentNode != bookmarkEnd)
                    {
                        var nextNode = currentNode.NextSibling;
                        currentNode.Remove();
                        currentNode = nextNode;
                    }
                    
                    // Insert new text
                    var builder = new DocumentBuilder(doc);
                    builder.MoveTo(bookmarkRange);
                    builder.Write(newText);
                    
                    changes.Add($"書籤內容已更新");
                }
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException($"無法更新書籤內容: {ex.Message}", ex);
            }
        }
        
        if (changes.Count == 0)
        {
            return await Task.FromResult($"未進行任何變更。請提供 newName 或 newText 參數。");
        }
        
        doc.Save(outputPath);
        
        var result = $"成功編輯書籤 '{bookmarkName}'\n";
        result += $"原名稱: {oldName}\n";
        result += $"原內容: {oldText}\n";
        if (!string.IsNullOrEmpty(newName))
        {
            result += $"新名稱: {newName}\n";
        }
        if (!string.IsNullOrEmpty(newText))
        {
            result += $"新內容: {newText}\n";
        }
        result += $"變更: {string.Join(", ", changes)}\n";
        result += $"輸出: {outputPath}";
        
        return await Task.FromResult(result);
    }
}

