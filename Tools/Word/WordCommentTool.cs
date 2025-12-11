using System.Text.Json.Nodes;
using Aspose.Words;
using AsposeMcpServer.Core;

namespace AsposeMcpServer.Tools;

/// <summary>
/// Unified tool for managing Word comments (add, edit, delete, get, reply)
/// Merges: WordAddCommentTool, WordDeleteCommentTool, WordGetCommentsTool, WordReplyCommentTool
/// </summary>
public class WordCommentTool : IAsposeTool
{
    public string Description => "Manage Word comments: add, delete, get all, or reply to a comment";

    public object InputSchema => new
    {
        type = "object",
        properties = new
        {
            operation = new
            {
                type = "string",
                description = "Operation to perform: 'add', 'delete', 'get', 'reply'",
                @enum = new[] { "add", "delete", "get", "reply" }
            },
            path = new
            {
                type = "string",
                description = "Document file path"
            },
            outputPath = new
            {
                type = "string",
                description = "Output file path (if not provided, overwrites input, for add/delete/reply operations)"
            },
            text = new
            {
                type = "string",
                description = "Comment text content (required for add operation)"
            },
            author = new
            {
                type = "string",
                description = "Comment author name (optional, defaults to 'Comment Author', for add/reply operations)"
            },
            paragraphIndex = new
            {
                type = "number",
                description = "Paragraph index to attach comment to (0-based, optional, for add operation)"
            },
            startRunIndex = new
            {
                type = "number",
                description = "Start run index within the paragraph (0-based, optional, for add operation)"
            },
            endRunIndex = new
            {
                type = "number",
                description = "End run index within the paragraph (0-based, optional, for add operation)"
            },
            commentIndex = new
            {
                type = "number",
                description = "Comment index (0-based, required for delete/reply operations)"
            },
            replyText = new
            {
                type = "string",
                description = "Reply text content (required for reply operation)"
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
            "add" => await AddCommentAsync(arguments, path),
            "delete" => await DeleteCommentAsync(arguments, path),
            "get" => await GetCommentsAsync(arguments, path),
            "reply" => await ReplyCommentAsync(arguments, path),
            _ => throw new ArgumentException($"Unknown operation: {operation}")
        };
    }

    private async Task<string> AddCommentAsync(JsonObject? arguments, string path)
    {
        var outputPath = arguments?["outputPath"]?.GetValue<string>() ?? path;
        var text = arguments?["text"]?.GetValue<string>() ?? throw new ArgumentException("text is required for add operation");
        var author = arguments?["author"]?.GetValue<string>() ?? "Comment Author";
        var paragraphIndex = arguments?["paragraphIndex"]?.GetValue<int?>();
        var startRunIndex = arguments?["startRunIndex"]?.GetValue<int?>();
        var endRunIndex = arguments?["endRunIndex"]?.GetValue<int?>();

        SecurityHelper.ValidateFilePath(outputPath, "outputPath");

        var doc = new Document(path);
        var builder = new DocumentBuilder(doc);
        var paragraphs = doc.GetChildNodes(NodeType.Paragraph, true);
        
        Paragraph? targetPara = null;
        
        // Determine target paragraph
        if (paragraphIndex.HasValue)
        {
            if (paragraphIndex.Value < 0 || paragraphIndex.Value >= paragraphs.Count)
            {
                throw new ArgumentException($"段落索引 {paragraphIndex.Value} 超出範圍 (文檔共有 {paragraphs.Count} 個段落)");
            }
            targetPara = paragraphs[paragraphIndex.Value] as Paragraph;
            if (targetPara == null)
            {
                throw new ArgumentException($"無法找到索引 {paragraphIndex.Value} 的段落");
            }
        }
        else
        {
            // Default: use last paragraph
            if (paragraphs.Count > 0)
            {
                targetPara = paragraphs[paragraphs.Count - 1] as Paragraph;
                if (targetPara == null)
                {
                    throw new InvalidOperationException("無法找到有效的段落");
                }
            }
            else
            {
                throw new InvalidOperationException("文檔中沒有段落可以添加註解");
            }
        }
        
        // Create comment
        var comment = new Comment(doc)
        {
            Author = author,
            Initial = author.Length >= 2 ? author.Substring(0, 2).ToUpper() : author.ToUpper(),
            DateTime = System.DateTime.Now
        };
        
        // Add comment text
        var commentPara = new Paragraph(doc);
        commentPara.Runs.Add(new Run(doc, text));
        comment.Paragraphs.Add(commentPara);
        
        // Determine comment range and insert comment
        if (startRunIndex.HasValue && endRunIndex.HasValue)
        {
            // Comment on specific runs
            var runs = targetPara.GetChildNodes(NodeType.Run, false)!;
            if (runs.Count == 0)
            {
                throw new InvalidOperationException("無法獲取 Run 節點");
            }
            if (startRunIndex.Value < 0 || startRunIndex.Value >= runs.Count ||
                endRunIndex.Value < 0 || endRunIndex.Value >= runs.Count ||
                startRunIndex.Value > endRunIndex.Value)
            {
                throw new ArgumentException($"Run 索引超出範圍 (段落共有 {runs.Count} 個 Run)");
            }
            
            var startRun = runs[startRunIndex.Value] as Run;
            var endRun = runs[endRunIndex.Value] as Run;
            
            if (startRun != null && endRun != null && startRun.ParentNode != null && endRun.ParentNode != null)
            {
                var rangeStart = new CommentRangeStart(doc, comment.Id);
                var rangeEnd = new CommentRangeEnd(doc, comment.Id);
                
                startRun.ParentNode.InsertBefore(rangeStart, startRun);
                endRun.ParentNode.InsertAfter(rangeEnd, endRun);
            }
        }
        else if (startRunIndex.HasValue)
        {
            // Comment on single run
            var runs = targetPara.GetChildNodes(NodeType.Run, false)!;
            if (runs.Count == 0)
            {
                throw new InvalidOperationException("無法獲取 Run 節點");
            }
            if (startRunIndex.Value < 0 || startRunIndex.Value >= runs.Count)
            {
                throw new ArgumentException($"Run 索引超出範圍 (段落共有 {runs.Count} 個 Run)");
            }
            
            var run = runs[startRunIndex.Value] as Run;
            if (run != null && run.ParentNode != null)
            {
                var rangeStart = new CommentRangeStart(doc, comment.Id);
                var rangeEnd = new CommentRangeEnd(doc, comment.Id);
                
                run.ParentNode.InsertBefore(rangeStart, run);
                run.ParentNode.InsertAfter(rangeEnd, run);
            }
        }
        else
        {
            // Comment on entire paragraph - insert at end of paragraph
            var rangeStart = new CommentRangeStart(doc, comment.Id);
            var rangeEnd = new CommentRangeEnd(doc, comment.Id);
            
            targetPara?.AppendChild(rangeStart);
            targetPara?.AppendChild(rangeEnd);
        }
        
        // Insert comment into document
        if (doc.FirstSection?.Body != null)
        {
            doc.FirstSection.Body.AppendChild(comment);
        }
        
        doc.Save(outputPath);
        
        var result = $"成功添加註解\n";
        result += $"作者: {author}\n";
        result += $"內容: {text}\n";
        if (paragraphIndex.HasValue)
        {
            result += $"段落索引: {paragraphIndex.Value}\n";
        }
        if (startRunIndex.HasValue)
        {
            result += $"Run 範圍: {startRunIndex.Value}";
            if (endRunIndex.HasValue)
            {
                result += $" - {endRunIndex.Value}";
            }
            result += "\n";
        }
        result += $"輸出: {outputPath}";
        
        return await Task.FromResult(result);
    }

    private async Task<string> DeleteCommentAsync(JsonObject? arguments, string path)
    {
        var outputPath = arguments?["outputPath"]?.GetValue<string>() ?? path;
        var commentIndex = arguments?["commentIndex"]?.GetValue<int>() ?? throw new ArgumentException("commentIndex is required for delete operation");

        SecurityHelper.ValidateFilePath(outputPath, "outputPath");

        var doc = new Document(path);
        
        // Get all comments
        var comments = doc.GetChildNodes(NodeType.Comment, true);
        
        if (commentIndex < 0 || commentIndex >= comments.Count)
        {
            throw new ArgumentException($"註解索引 {commentIndex} 超出範圍 (文檔共有 {comments.Count} 個註解)");
        }
        
        var commentToDelete = comments[commentIndex] as Comment;
        if (commentToDelete == null)
        {
            throw new InvalidOperationException($"無法找到索引 {commentIndex} 的註解");
        }
        
        // Get comment info before deletion
        string author = commentToDelete.Author;
        string commentText = commentToDelete.GetText().Trim();
        string preview = commentText.Length > 50 ? commentText.Substring(0, 50) + "..." : commentText;
        
        // Remove comment range markers if they exist
        var rangeStarts = doc.GetChildNodes(NodeType.CommentRangeStart, true);
        var rangeEnds = doc.GetChildNodes(NodeType.CommentRangeEnd, true);
        
        foreach (CommentRangeStart rangeStart in rangeStarts)
        {
            if (rangeStart.Id == commentToDelete.Id)
            {
                rangeStart.Remove();
            }
        }
        
        foreach (CommentRangeEnd rangeEnd in rangeEnds)
        {
            if (rangeEnd.Id == commentToDelete.Id)
            {
                rangeEnd.Remove();
            }
        }
        
        // Delete the comment
        commentToDelete.Remove();
        
        doc.Save(outputPath);
        
        // Count remaining comments
        int remainingCount = doc.GetChildNodes(NodeType.Comment, true).Count;
        
        var result = $"成功刪除註解 #{commentIndex}\n";
        result += $"作者: {author}\n";
        result += $"內容預覽: {preview}\n";
        result += $"文檔剩餘註解數: {remainingCount}\n";
        result += $"輸出: {outputPath}";
        
        return await Task.FromResult(result);
    }

    private async Task<string> GetCommentsAsync(JsonObject? arguments, string path)
    {
        var doc = new Document(path);
        
        // Get all comments
        var comments = doc.GetChildNodes(NodeType.Comment, true);
        
        if (comments.Count == 0)
        {
            return await Task.FromResult("文檔中沒有找到註解");
        }
        
        var result = new System.Text.StringBuilder();
        result.AppendLine($"找到 {comments.Count} 個註解：\n");
        
        int index = 0;
        foreach (Comment comment in comments)
        {
            result.AppendLine($"註解 #{index}:");
            result.AppendLine($"  作者: {comment.Author}");
            result.AppendLine($"  初始: {comment.Initial}");
            result.AppendLine($"  日期: {comment.DateTime:yyyy-MM-dd HH:mm:ss}");
            
            // Get comment text
            string commentText = comment.GetText().Trim();
            if (commentText.Length > 100)
            {
                commentText = commentText.Substring(0, 100) + "...";
            }
            result.AppendLine($"  內容: {commentText}");
            
            // Get commented text range if available
            var rangeStarts = doc.GetChildNodes(NodeType.CommentRangeStart, true);
            
            foreach (CommentRangeStart rangeStart in rangeStarts)
            {
                if (rangeStart.Id == comment.Id)
                {
                    result.AppendLine($"  範圍: 已標記文字");
                    break;
                }
            }
            
            result.AppendLine();
            index++;
        }
        
        return await Task.FromResult(result.ToString().TrimEnd());
    }

    private async Task<string> ReplyCommentAsync(JsonObject? arguments, string path)
    {
        var outputPath = arguments?["outputPath"]?.GetValue<string>() ?? path;
        var commentIndex = arguments?["commentIndex"]?.GetValue<int>() ?? throw new ArgumentException("commentIndex is required for reply operation");
        var replyText = arguments?["replyText"]?.GetValue<string>() ?? throw new ArgumentException("replyText is required for reply operation");
        var author = arguments?["author"]?.GetValue<string>() ?? "Reply Author";

        SecurityHelper.ValidateFilePath(outputPath, "outputPath");

        var doc = new Document(path);
        
        // Get all comments
        var comments = doc.GetChildNodes(NodeType.Comment, true);
        
        if (commentIndex < 0 || commentIndex >= comments.Count)
        {
            throw new ArgumentException($"註解索引 {commentIndex} 超出範圍 (文檔共有 {comments.Count} 個註解)");
        }
        
        var parentComment = comments[commentIndex] as Comment;
        if (parentComment == null)
        {
            throw new InvalidOperationException($"無法找到索引 {commentIndex} 的註解");
        }
        
        // Create reply comment
        var replyComment = new Comment(doc)
        {
            Author = author,
            Initial = author.Substring(0, Math.Min(2, author.Length)).ToUpper()
        };
        
        replyComment.Paragraphs.Add(new Paragraph(doc));
        replyComment.FirstParagraph.Runs.Add(new Run(doc, replyText));
        
        // Add reply to parent comment
        // In Aspose.Words, replies are added as child comments
        parentComment.AppendChild(replyComment);
        
        doc.Save(outputPath);
        
        var result = $"成功回覆註解 #{commentIndex}\n";
        result += $"原註解作者: {parentComment.Author}\n";
        result += $"回覆作者: {author}\n";
        result += $"回覆內容: {replyText}\n";
        result += $"輸出: {outputPath}";
        
        return await Task.FromResult(result);
    }
}

