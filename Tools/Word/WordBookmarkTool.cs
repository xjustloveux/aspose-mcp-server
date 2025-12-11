using System.Text.Json.Nodes;
using Aspose.Words;
using AsposeMcpServer.Core;

namespace AsposeMcpServer.Tools;

/// <summary>
/// Unified tool for managing Word bookmarks (add, edit, delete, get, goto)
/// Merges: WordAddBookmarkTool, WordEditBookmarkTool, WordDeleteBookmarkTool, WordGetBookmarksTool, WordGotoBookmarkTool
/// </summary>
public class WordBookmarkTool : IAsposeTool
{
    public string Description => "Manage Word bookmarks: add, edit, delete, get all, or goto (get location)";

    public object InputSchema => new
    {
        type = "object",
        properties = new
        {
            operation = new
            {
                type = "string",
                description = "Operation to perform: 'add', 'edit', 'delete', 'get', 'goto'",
                @enum = new[] { "add", "edit", "delete", "get", "goto" }
            },
            path = new
            {
                type = "string",
                description = "Document file path"
            },
            outputPath = new
            {
                type = "string",
                description = "Output file path (if not provided, overwrites input, for add/edit/delete operations)"
            },
            name = new
            {
                type = "string",
                description = "Bookmark name (required for add/edit/delete/goto operations)"
            },
            text = new
            {
                type = "string",
                description = "Text content to bookmark (optional, for add operation)"
            },
            paragraphIndex = new
            {
                type = "number",
                description = "Paragraph index to insert bookmark at (0-based, optional, for add operation)"
            },
            newName = new
            {
                type = "string",
                description = "New bookmark name (optional, for edit operation)"
            },
            newText = new
            {
                type = "string",
                description = "New text content for the bookmark (optional, for edit operation)"
            },
            keepText = new
            {
                type = "boolean",
                description = "Keep bookmark text content when deleting (default: true, for delete operation)"
            }
        },
        required = new[] { "operation", "path" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var operation = arguments?["operation"]?.GetValue<string>() ?? throw new ArgumentException("operation is required");
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");

        SecurityHelper.ValidateFilePath(path, "path");

        return operation.ToLower() switch
        {
            "add" => await AddBookmarkAsync(arguments, path),
            "edit" => await EditBookmarkAsync(arguments, path),
            "delete" => await DeleteBookmarkAsync(arguments, path),
            "get" => await GetBookmarksAsync(arguments, path),
            "goto" => await GotoBookmarkAsync(arguments, path),
            _ => throw new ArgumentException($"Unknown operation: {operation}")
        };
    }

    private async Task<string> AddBookmarkAsync(JsonObject? arguments, string path)
    {
        var outputPath = arguments?["outputPath"]?.GetValue<string>() ?? path;
        var name = arguments?["name"]?.GetValue<string>() ?? throw new ArgumentException("name is required for add operation");
        var text = arguments?["text"]?.GetValue<string>();
        var paragraphIndex = arguments?["paragraphIndex"]?.GetValue<int?>();

        SecurityHelper.ValidateFilePath(outputPath, "outputPath");

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

    private async Task<string> EditBookmarkAsync(JsonObject? arguments, string path)
    {
        var outputPath = arguments?["outputPath"]?.GetValue<string>() ?? path;
        var bookmarkName = arguments?["name"]?.GetValue<string>() ?? throw new ArgumentException("name is required for edit operation");
        var newName = arguments?["newName"]?.GetValue<string>();
        var newText = arguments?["newText"]?.GetValue<string>();

        SecurityHelper.ValidateFilePath(outputPath, "outputPath");

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
                throw new InvalidOperationException($"無法重新命名書籤: {ex.Message}", ex);
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

    private async Task<string> DeleteBookmarkAsync(JsonObject? arguments, string path)
    {
        var outputPath = arguments?["outputPath"]?.GetValue<string>() ?? path;
        var name = arguments?["name"]?.GetValue<string>() ?? throw new ArgumentException("name is required for delete operation");
        var keepText = arguments?["keepText"]?.GetValue<bool>() ?? true;

        SecurityHelper.ValidateFilePath(outputPath, "outputPath");

        var doc = new Document(path);
        
        // Get bookmark
        var bookmark = doc.Range.Bookmarks[name];
        if (bookmark == null)
        {
            throw new ArgumentException($"找不到書籤 '{name}'，可用書籤請使用 get 操作查看");
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

    private async Task<string> GetBookmarksAsync(JsonObject? arguments, string path)
    {
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

    private async Task<string> GotoBookmarkAsync(JsonObject? arguments, string path)
    {
        var name = arguments?["name"]?.GetValue<string>() ?? throw new ArgumentException("name is required for goto operation");

        var doc = new Document(path);
        
        // Get bookmark
        var bookmark = doc.Range.Bookmarks[name];
        if (bookmark == null)
        {
            throw new ArgumentException($"找不到書籤 '{name}'，可用書籤請使用 get 操作查看");
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

