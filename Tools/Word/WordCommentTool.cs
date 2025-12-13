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
    public string Description => @"Manage Word comments. Supports 4 operations: add, delete, get, reply.

Usage examples:
- Add comment: word_comment(operation='add', path='doc.docx', text='This is a comment', author='Author Name')
- Delete comment: word_comment(operation='delete', path='doc.docx', commentIndex=0)
- Get all comments: word_comment(operation='get', path='doc.docx')
- Reply to comment: word_comment(operation='reply', path='doc.docx', commentIndex=0, text='This is a reply')";

    public object InputSchema => new
    {
        type = "object",
        properties = new
        {
            operation = new
            {
                type = "string",
                description = @"Operation to perform.
- 'add': Add a new comment (required params: path, text)
- 'delete': Delete a comment (required params: path, commentIndex)
- 'get': Get all comments (required params: path)
- 'reply': Reply to a comment (required params: path, commentIndex, text)",
                @enum = new[] { "add", "delete", "get", "reply" }
            },
            path = new
            {
                type = "string",
                description = "Document file path (required for all operations)"
            },
            outputPath = new
            {
                type = "string",
                description = "Output file path (if not provided, overwrites input, for add/delete/reply operations)"
            },
            text = new
            {
                type = "string",
                description = "Comment text content (required for add and reply operations)"
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
        // IMPORTANT: Only get paragraphs from document Body, not from Comment objects
        // Using GetChildNodes(NodeType.Paragraph, true) would recursively include paragraphs
        // inside Comment objects, which causes index calculation errors after replies are added
        var paragraphs = new List<Paragraph>();
        foreach (Section section in doc.Sections)
        {
            var bodyParagraphs = section.Body.GetChildNodes(NodeType.Paragraph, false).Cast<Paragraph>().ToList();
            paragraphs.AddRange(bodyParagraphs);
        }
        
        Paragraph? targetPara = null;
        Run? startRun = null;
        Run? endRun = null;
        
        // Determine target paragraph and runs
        if (paragraphIndex.HasValue)
        {
            // Support -1 for document end (like other tools)
            if (paragraphIndex.Value == -1)
            {
                // Use last paragraph
                if (paragraphs.Count > 0)
                {
                    targetPara = paragraphs[paragraphs.Count - 1] as Paragraph;
                }
                else
                {
                    throw new InvalidOperationException("文檔中沒有段落可以添加註解");
                }
            }
            else if (paragraphIndex.Value < 0 || paragraphIndex.Value >= paragraphs.Count)
            {
                throw new ArgumentException($"段落索引 {paragraphIndex.Value} 超出範圍 (文檔共有 {paragraphs.Count} 個段落，有效索引: 0-{paragraphs.Count - 1} 或 -1 表示最後一個段落)");
            }
            else
            {
                targetPara = paragraphs[paragraphIndex.Value] as Paragraph;
                if (targetPara == null)
                {
                    throw new ArgumentException($"無法找到索引 {paragraphIndex.Value} 的段落");
                }
            }
        }
        else
        {
            var builder = new DocumentBuilder(doc);
            builder.MoveToDocumentEnd();
            
            // Create a new paragraph for this comment
            var newPara = new Paragraph(doc);
            var newRun = new Run(doc, " "); // Add a space so paragraph is not empty
            newPara.AppendChild(newRun);
            doc.LastSection.Body.AppendChild(newPara);
            
            targetPara = newPara;
            
            // Update paragraphs list to include the new paragraph
            paragraphs.Clear();
            foreach (Section section in doc.Sections)
            {
                var bodyParagraphs = section.Body.GetChildNodes(NodeType.Paragraph, false).Cast<Paragraph>().ToList();
                paragraphs.AddRange(bodyParagraphs);
            }
        }
        
        // Determine which runs to comment on
        if (targetPara == null)
        {
            throw new InvalidOperationException("無法確定目標段落");
        }
        var runs = targetPara.GetChildNodes(NodeType.Run, false);
        if (runs == null || runs.Count == 0)
        {
            // If paragraph has no runs, create a placeholder run
            var placeholderRun = new Run(doc, " ");
            targetPara.AppendChild(placeholderRun);
            startRun = placeholderRun;
            endRun = placeholderRun;
        }
        else if (startRunIndex.HasValue && endRunIndex.HasValue)
        {
            // Comment on specific run range
            if (startRunIndex.Value < 0 || startRunIndex.Value >= runs.Count ||
                endRunIndex.Value < 0 || endRunIndex.Value >= runs.Count ||
                startRunIndex.Value > endRunIndex.Value)
            {
                throw new ArgumentException($"Run 索引超出範圍 (段落共有 {runs.Count} 個 Run)");
            }
            startRun = runs[startRunIndex.Value] as Run;
            endRun = runs[endRunIndex.Value] as Run;
        }
        else if (startRunIndex.HasValue)
        {
            // Comment on single run
            if (startRunIndex.Value < 0 || startRunIndex.Value >= runs.Count)
            {
                throw new ArgumentException($"Run 索引超出範圍 (段落共有 {runs.Count} 個 Run)");
            }
            startRun = runs[startRunIndex.Value] as Run;
            endRun = startRun;
        }
        else
        {
            // Comment on entire paragraph - use first and last runs
            startRun = runs[0] as Run;
            endRun = runs[runs.Count - 1] as Run;
        }
        
        if (startRun == null || endRun == null)
        {
            throw new InvalidOperationException("無法確定評論範圍");
        }
        
        // Get the paragraph containing the runs
        Paragraph? para = startRun.ParentNode as Paragraph;
        if (para == null)
        {
            para = startRun.GetAncestor(NodeType.Paragraph) as Paragraph;
        }
        
        if (para == null)
        {
            throw new InvalidOperationException("無法找到包含 Run 的段落節點");
        }
        
        var comment = new Comment(doc, author, author.Length >= 2 ? author.Substring(0, 2).ToUpper() : author.ToUpper(), System.DateTime.Now);
        comment.Paragraphs.Add(new Paragraph(doc));
        comment.FirstParagraph.AppendChild(new Run(doc, text));
        
        var rangeStart = new CommentRangeStart(doc, comment.Id);
        var rangeEnd = new CommentRangeEnd(doc, comment.Id);
        
        try
        {
            // Get the paragraph containing startRun (use the one we already found)
            var startPara = para; // Use the paragraph we found earlier
            
            // Insert CommentRangeStart before startRun
            // CommentRangeStart must be inserted directly into the paragraph, not nested
            // Ensure startRun is a direct child of startPara
            if (startRun.ParentNode != startPara)
            {
                // If startRun is not a direct child, we need to find its actual parent
                // and ensure we're inserting into the correct paragraph
                var actualParent = startRun.ParentNode;
                if (actualParent != null && actualParent.NodeType == NodeType.Paragraph)
                {
                    var parentPara = actualParent as Paragraph;
                    if (parentPara != null)
                    {
                        startPara = parentPara;
                    }
                }
            }
            
            // Now insert CommentRangeStart before startRun
            if (startRun.ParentNode == startPara)
            {
                startPara.InsertBefore(rangeStart, startRun);
            }
            else
            {
                // Fallback: insert at the beginning of the paragraph
                startPara.InsertBefore(rangeStart, startPara.FirstChild);
            }
            
            // Get the paragraph containing endRun
            Paragraph? endPara = endRun.ParentNode as Paragraph;
            if (endPara == null)
            {
                endPara = endRun.GetAncestor(NodeType.Paragraph) as Paragraph;
            }
            
            if (endPara == null)
            {
                throw new InvalidOperationException($"無法找到包含 endRun 的段落");
            }
            
            // Insert CommentRangeEnd after endRun
            if (endRun.ParentNode == endPara)
            {
                var nextSibling = endRun.NextSibling;
                if (nextSibling != null)
                {
                    endPara.InsertBefore(rangeEnd, nextSibling);
                }
                else
                {
                    endPara.AppendChild(rangeEnd);
                }
            }
            else
            {
                // Insert at the end of the paragraph
                endPara.AppendChild(rangeEnd);
            }
            
            // Comment objects are linked to CommentRangeStart/End via Id
            // Insert Comment object into the paragraph containing the comment range
            // Note: The Comment object itself needs to be inserted into the paragraph structure,
            // but its content (comment.Paragraphs) should not appear in the document body
            if (endPara != null)
            {
                // Find CommentRangeEnd position and insert Comment after it in the paragraph
                var rangeEndNode = endPara.GetChildNodes(NodeType.CommentRangeEnd, false)
                    .Cast<CommentRangeEnd>()
                    .FirstOrDefault(re => re.Id == comment.Id);
                
                if (rangeEndNode != null)
                {
                    // Insert Comment after CommentRangeEnd in the paragraph
                    // Check if comment is already in the document structure
                    if (comment.ParentNode == null)
                    {
                        endPara.InsertAfter(comment, rangeEndNode);
                    }
                    else
                    {
                        // Comment is already in the document structure, verify it's in the correct location
                        if (comment.ParentNode != endPara)
                        {
                            // Comment is in wrong location, remove and reinsert
                            comment.Remove();
                            endPara.InsertAfter(comment, rangeEndNode);
                        }
                    }
                }
                else
                {
                    // Fallback: append to paragraph end
                    // Check if comment is already in the document structure
                    if (comment.ParentNode == null)
                    {
                        endPara.AppendChild(comment);
                    }
                    else
                    {
                        // Comment is already in the document structure, verify it's in the correct location
                        if (comment.ParentNode != endPara)
                        {
                            // Comment is in wrong location, remove and reinsert
                            comment.Remove();
                            endPara.AppendChild(comment);
                        }
                    }
                }
            }
            else
            {
                throw new InvalidOperationException($"無法找到目標段落來插入 Comment 對象");
            }
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException($"插入評論範圍標記時發生錯誤: {ex.Message}", ex);
        }
        
        doc.EnsureMinimum();
        
        // Verify that the comment was actually added to the document
        // This is important because in some cases (e.g., when there are existing replies),
        // the comment insertion might fail silently
        var allCommentsAfter = doc.GetChildNodes(NodeType.Comment, true);
        bool commentFound = false;
        foreach (Comment existingComment in allCommentsAfter)
        {
            if (existingComment.Id == comment.Id)
            {
                commentFound = true;
                break;
            }
        }
        
        if (!commentFound)
        {
            throw new InvalidOperationException($"評論添加失敗：評論對象未能成功插入到文檔中。這可能發生在已有回復的情況下。");
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
        var allComments = doc.GetChildNodes(NodeType.Comment, true).Cast<Comment>().ToList();
        
        // Build a set of all reply comment IDs (comments that are replies to other comments)
        var replyCommentIds = new HashSet<int>();
        foreach (Comment comment in allComments)
        {
            if (comment.Replies != null && comment.Replies.Count > 0)
            {
                foreach (Comment reply in comment.Replies)
                {
                    replyCommentIds.Add(reply.Id);
                }
            }
        }
        
        // Filter to get only top-level comments:
        // 1. Comments without an ancestor (Ancestor == null)
        // 2. Comments that are not in any other comment's Replies collection
        var topLevelComments = new List<Comment>();
        foreach (Comment comment in allComments)
        {
            // A comment is top-level if:
            // - It has no ancestor (Ancestor == null)
            // - It's not in the replyCommentIds set (not a reply to another comment)
            if (comment.Ancestor == null && !replyCommentIds.Contains(comment.Id))
            {
                if (!topLevelComments.Any(c => c.Id == comment.Id))
                {
                    topLevelComments.Add(comment);
                }
            }
        }
        
        // Also check comments from CommentRangeStart nodes
        var rangeStarts = doc.GetChildNodes(NodeType.CommentRangeStart, true);
        foreach (CommentRangeStart rangeStart in rangeStarts)
        {
            var commentById = allComments.FirstOrDefault(c => c.Id == rangeStart.Id);
            if (commentById != null && 
                commentById.Ancestor == null && 
                !replyCommentIds.Contains(commentById.Id) &&
                !topLevelComments.Any(c => c.Id == commentById.Id))
            {
                topLevelComments.Add(commentById);
            }
        }
        
        if (topLevelComments.Count == 0)
        {
            return await Task.FromResult("文檔中沒有找到註解");
        }
        
        var result = new System.Text.StringBuilder();
        result.AppendLine($"找到 {topLevelComments.Count} 個頂層註解：\n");
        
        int index = 0;
        foreach (Comment comment in topLevelComments)
        {
            AppendCommentInfo(result, comment, doc, index, 0);
            index++;
        }
        
        return await Task.FromResult(result.ToString().TrimEnd());
    }
    
    private void AppendCommentInfo(System.Text.StringBuilder result, Comment comment, Document doc, int index, int indentLevel)
    {
        string indent = new string(' ', indentLevel * 2);
        string prefix = indentLevel == 0 ? $"註解 #{index}" : $"  └─ 回復";
        
        result.AppendLine($"{indent}{prefix}:");
        result.AppendLine($"{indent}  作者: {comment.Author}");
        result.AppendLine($"{indent}  初始: {comment.Initial}");
        result.AppendLine($"{indent}  日期: {comment.DateTime:yyyy-MM-dd HH:mm:ss}");
        
        // Get comment text
        string commentText = comment.GetText().Trim();
        if (commentText.Length > 100)
        {
            commentText = commentText.Substring(0, 100) + "...";
        }
        result.AppendLine($"{indent}  內容: {commentText}");
        
        // Get commented text range if available
        var commentRangeStarts = doc.GetChildNodes(NodeType.CommentRangeStart, true);
        bool hasRange = false;
        foreach (CommentRangeStart rangeStart in commentRangeStarts)
        {
            if (rangeStart.Id == comment.Id)
            {
                result.AppendLine($"{indent}  範圍: 已標記文字");
                hasRange = true;
                break;
            }
        }
        if (!hasRange)
        {
            result.AppendLine($"{indent}  範圍: 未找到範圍標記");
        }
        
        // Display replies if any
        if (comment.Replies != null && comment.Replies.Count > 0)
        {
            result.AppendLine($"{indent}  回復數: {comment.Replies.Count}");
            foreach (Comment reply in comment.Replies)
            {
                AppendCommentInfo(result, reply, doc, -1, indentLevel + 1);
            }
        }
        
        result.AppendLine();
    }

    private async Task<string> ReplyCommentAsync(JsonObject? arguments, string path)
    {
        var outputPath = arguments?["outputPath"]?.GetValue<string>() ?? path;
        var commentIndex = arguments?["commentIndex"]?.GetValue<int>() ?? throw new ArgumentException("commentIndex is required for reply operation");
        // Accept both replyText and text for compatibility
        var replyText = arguments?["replyText"]?.GetValue<string>() ?? 
                        arguments?["text"]?.GetValue<string>() ?? 
                        throw new ArgumentException("replyText or text is required for reply operation");
        var author = arguments?["author"]?.GetValue<string>() ?? "Reply Author";

        SecurityHelper.ValidateFilePath(outputPath, "outputPath");

        var doc = new Document(path);
        
        // Get all comments and filter to top-level comments only (same logic as GetCommentsAsync)
        // This ensures commentIndex refers to top-level comments, not replies
        var allComments = doc.GetChildNodes(NodeType.Comment, true).Cast<Comment>().ToList();
        
        // Build a set of all reply comment IDs (comments that are replies to other comments)
        var replyCommentIds = new HashSet<int>();
        foreach (Comment comment in allComments)
        {
            if (comment.Replies != null && comment.Replies.Count > 0)
            {
                foreach (Comment reply in comment.Replies)
                {
                    replyCommentIds.Add(reply.Id);
                }
            }
        }
        
        // Filter to get only top-level comments:
        // 1. Comments without an ancestor (Ancestor == null)
        // 2. Comments that are not in any other comment's Replies collection
        var topLevelComments = new List<Comment>();
        foreach (Comment comment in allComments)
        {
            if (comment.Ancestor == null && !replyCommentIds.Contains(comment.Id))
            {
                if (!topLevelComments.Any(c => c.Id == comment.Id))
                {
                    topLevelComments.Add(comment);
                }
            }
        }
        
        // Also check comments from CommentRangeStart nodes
        var rangeStarts = doc.GetChildNodes(NodeType.CommentRangeStart, true);
        foreach (CommentRangeStart rangeStart in rangeStarts)
        {
            var commentById = allComments.FirstOrDefault(c => c.Id == rangeStart.Id);
            if (commentById != null && 
                commentById.Ancestor == null && 
                !replyCommentIds.Contains(commentById.Id) &&
                !topLevelComments.Any(c => c.Id == commentById.Id))
            {
                topLevelComments.Add(commentById);
            }
        }
        
        // Validate commentIndex against top-level comments only
        if (commentIndex < 0 || commentIndex >= topLevelComments.Count)
        {
            throw new ArgumentException($"註解索引 {commentIndex} 超出範圍 (文檔共有 {topLevelComments.Count} 個頂層註解)");
        }
        
        var parentComment = topLevelComments[commentIndex];
        
        // Check if parentComment is actually a reply (should not happen, but safety check)
        if (parentComment.Ancestor != null)
        {
            throw new InvalidOperationException($"註解索引 {commentIndex} 指向一個回復，無法對回復添加回復。請使用頂層註解的索引。");
        }
        
        // Use AddReply() method - the correct Aspose.Words API for adding replies
        // AddReply() creates the reply comment and adds it to the parent comment's Replies collection
        // It does NOT insert the reply content into the document body
        var initial = author.Length >= 2 ? author.Substring(0, 2).ToUpper() : author.ToUpper();
        Comment replyComment = parentComment.AddReply(author, initial, System.DateTime.Now, replyText);
        doc.Save(outputPath);
        
        var result = $"成功回覆註解 #{commentIndex}\n";
        result += $"原註解作者: {parentComment.Author}\n";
        result += $"回覆作者: {author}\n";
        result += $"回覆內容: {replyText}\n";
        result += $"輸出: {outputPath}";
        
        return await Task.FromResult(result);
    }
}

