using System.Text;
using System.Text.Json.Nodes;
using Aspose.Words;
using AsposeMcpServer.Core;

namespace AsposeMcpServer.Tools.Word;

/// <summary>
///     Unified tool for managing Word comments (add, delete, get, reply)
///     Merges: WordAddCommentTool, WordDeleteCommentTool, WordGetCommentsTool, WordReplyCommentTool
/// </summary>
public class WordCommentTool : IAsposeTool
{
    /// <summary>
    ///     Gets the description of the tool and its usage examples
    /// </summary>
    public string Description => @"Manage Word comments. Supports 4 operations: add, delete, get, reply.

Usage examples:
- Add comment: word_comment(operation='add', path='doc.docx', text='This is a comment', author='Author Name')
- Delete comment: word_comment(operation='delete', path='doc.docx', commentIndex=0)
- Get all comments: word_comment(operation='get', path='doc.docx')
- Reply to comment: word_comment(operation='reply', path='doc.docx', commentIndex=0, text='This is a reply')";

    /// <summary>
    ///     Gets the JSON schema defining the input parameters for the tool
    /// </summary>
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

    /// <summary>
    ///     Executes the tool operation with the provided JSON arguments
    /// </summary>
    /// <param name="arguments">JSON arguments object containing operation parameters</param>
    /// <returns>Result message as a string</returns>
    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var operation = ArgumentHelper.GetString(arguments, "operation");
        var path = ArgumentHelper.GetAndValidatePath(arguments);

        SecurityHelper.ValidateFilePath(path, allowAbsolutePaths: true);

        return operation.ToLower() switch
        {
            "add" => await AddCommentAsync(arguments, path),
            "delete" => await DeleteCommentAsync(arguments, path),
            "get" => await GetCommentsAsync(arguments, path),
            "reply" => await ReplyCommentAsync(arguments, path),
            _ => throw new ArgumentException($"Unknown operation: {operation}")
        };
    }

    /// <summary>
    ///     Adds a new comment to the document
    /// </summary>
    /// <param name="arguments">
    ///     JSON arguments containing text, optional author, paragraphIndex, startRunIndex, endRunIndex,
    ///     outputPath
    /// </param>
    /// <param name="path">Word document file path</param>
    /// <returns>Success message with comment details</returns>
    private Task<string> AddCommentAsync(JsonObject? arguments, string path)
    {
        return Task.Run(() =>
        {
            var outputPath = ArgumentHelper.GetStringNullable(arguments, "outputPath") ?? path;
            var text = ArgumentHelper.GetString(arguments, "text");
            var author = ArgumentHelper.GetString(arguments, "author", "Comment Author");
            var paragraphIndex = ArgumentHelper.GetIntNullable(arguments, "paragraphIndex");
            var startRunIndex = ArgumentHelper.GetIntNullable(arguments, "startRunIndex");
            var endRunIndex = ArgumentHelper.GetIntNullable(arguments, "endRunIndex");

            SecurityHelper.ValidateFilePath(outputPath, "outputPath", true);

            var doc = new Document(path);
            // Only get paragraphs from document Body, not from Comment objects (to avoid index errors)
            var paragraphs = new List<Paragraph>();
            foreach (var section in doc.Sections.Cast<Section>())
            {
                var bodyParagraphs = section.Body.GetChildNodes(NodeType.Paragraph, false).Cast<Paragraph>().ToList();
                paragraphs.AddRange(bodyParagraphs);
            }

            Paragraph? targetPara;
            Run? startRun;
            Run? endRun;

            // Determine target paragraph and runs
            if (paragraphIndex.HasValue)
            {
                // Support -1 for document end (like other tools)
                if (paragraphIndex.Value == -1)
                {
                    // Use last paragraph
                    if (paragraphs.Count > 0)
                        targetPara = paragraphs[^1];
                    else
                        throw new InvalidOperationException("Document has no paragraphs to add comment to");
                }
                else if (paragraphIndex.Value < 0 || paragraphIndex.Value >= paragraphs.Count)
                {
                    throw new ArgumentException(
                        $"Paragraph index {paragraphIndex.Value} is out of range (document has {paragraphs.Count} paragraphs, valid index: 0-{paragraphs.Count - 1} or -1 for last paragraph)");
                }
                else
                {
                    targetPara = paragraphs[paragraphIndex.Value];
                    if (targetPara == null)
                        throw new ArgumentException($"Unable to find paragraph at index {paragraphIndex.Value}");
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
                foreach (var section in doc.Sections.Cast<Section>())
                {
                    var bodyParagraphs = section.Body.GetChildNodes(NodeType.Paragraph, false).Cast<Paragraph>()
                        .ToList();
                    paragraphs.AddRange(bodyParagraphs);
                }
            }

            // Determine which runs to comment on
            if (targetPara == null) throw new InvalidOperationException("Unable to determine target paragraph");
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
                if (startRunIndex.Value < 0 || startRunIndex.Value >= runs.Count ||
                    endRunIndex.Value < 0 || endRunIndex.Value >= runs.Count ||
                    startRunIndex.Value > endRunIndex.Value)
                    throw new ArgumentException($"Run index is out of range (paragraph has {runs.Count} Runs)");
                startRun = runs[startRunIndex.Value] as Run;
                endRun = runs[endRunIndex.Value] as Run;
            }
            else if (startRunIndex.HasValue)
            {
                if (startRunIndex.Value < 0 || startRunIndex.Value >= runs.Count)
                    throw new ArgumentException($"Run index is out of range (paragraph has {runs.Count} Runs)");
                startRun = runs[startRunIndex.Value] as Run;
                endRun = startRun;
            }
            else
            {
                startRun = runs[0] as Run;
                endRun = runs[^1] as Run;
            }

            if (startRun == null || endRun == null)
                throw new InvalidOperationException("Unable to determine comment range");

            var para = startRun.ParentNode as Paragraph ?? startRun.GetAncestor(NodeType.Paragraph) as Paragraph;

            if (para == null) throw new InvalidOperationException("Unable to find paragraph node containing Run");

            var comment = new Comment(doc, author,
                author.Length >= 2 ? author.Substring(0, 2).ToUpper() : author.ToUpper(),
                DateTime.Now);
            comment.Paragraphs.Add(new Paragraph(doc));
            comment.FirstParagraph.AppendChild(new Run(doc, text));

            var rangeStart = new CommentRangeStart(doc, comment.Id);
            var rangeEnd = new CommentRangeEnd(doc, comment.Id);

            try
            {
                var startPara = para;

                if (startRun.ParentNode != startPara)
                {
                    var actualParent = startRun.ParentNode;
                    if (actualParent is Paragraph parentPara) startPara = parentPara;
                }

                startPara.InsertBefore(rangeStart, startRun.ParentNode == startPara ? startRun : startPara.FirstChild);

                var endPara = endRun.ParentNode as Paragraph ?? endRun.GetAncestor(NodeType.Paragraph) as Paragraph;

                if (endPara == null) throw new InvalidOperationException("Unable to find paragraph containing endRun");

                if (endRun.ParentNode == endPara)
                {
                    var nextSibling = endRun.NextSibling;
                    if (nextSibling != null)
                        endPara.InsertBefore(rangeEnd, nextSibling);
                    else
                        endPara.AppendChild(rangeEnd);
                }
                else
                {
                    endPara.AppendChild(rangeEnd);
                }

                // Comment objects are linked to CommentRangeStart/End via Id
                var rangeEndNode = endPara.GetChildNodes(NodeType.CommentRangeEnd, false)
                    .Cast<CommentRangeEnd>()
                    .FirstOrDefault(re => re.Id == comment.Id);

                if (rangeEndNode != null)
                {
                    if (comment.ParentNode == null)
                    {
                        endPara.InsertAfter(comment, rangeEndNode);
                    }
                    else if (comment.ParentNode != endPara)
                    {
                        comment.Remove();
                        endPara.InsertAfter(comment, rangeEndNode);
                    }
                }
                else
                {
                    if (comment.ParentNode == null)
                    {
                        endPara.AppendChild(comment);
                    }
                    else if (comment.ParentNode != endPara)
                    {
                        comment.Remove();
                        endPara.AppendChild(comment);
                    }
                }
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException(
                    $"Error occurred while inserting comment range markers: {ex.Message}",
                    ex);
            }

            doc.EnsureMinimum();

            // Verify that the comment was actually added to the document
            // This is important because in some cases (e.g., when there are existing replies),
            // the comment insertion might fail silently
            var allCommentsAfter = doc.GetChildNodes(NodeType.Comment, true).Cast<Comment>();
            var commentFound = false;
            foreach (var existingComment in allCommentsAfter)
                if (existingComment.Id == comment.Id)
                {
                    commentFound = true;
                    break;
                }

            if (!commentFound)
                throw new InvalidOperationException(
                    "Comment addition failed: Comment object was not successfully inserted into the document. This may occur when there are existing replies.");

            doc.Save(outputPath);

            var result = "Comment added successfully\n";
            result += $"Author: {author}\n";
            result += $"Content: {text}\n";
            if (paragraphIndex.HasValue) result += $"Paragraph index: {paragraphIndex.Value}\n";
            if (startRunIndex.HasValue)
            {
                result += $"Run range: {startRunIndex.Value}";
                if (endRunIndex.HasValue) result += $" - {endRunIndex.Value}";
                result += "\n";
            }

            result += $"Output: {outputPath}";

            return result;
        });
    }

    /// <summary>
    ///     Deletes a comment from the document
    /// </summary>
    /// <param name="arguments">JSON arguments containing commentIndex and optional outputPath</param>
    /// <param name="path">Word document file path</param>
    /// <returns>Success message with remaining comment count</returns>
    private Task<string> DeleteCommentAsync(JsonObject? arguments, string path)
    {
        return Task.Run(() =>
        {
            var outputPath = ArgumentHelper.GetStringNullable(arguments, "outputPath") ?? path;
            var commentIndex = ArgumentHelper.GetInt(arguments, "commentIndex");

            SecurityHelper.ValidateFilePath(outputPath, "outputPath", true);

            var doc = new Document(path);

            var comments = doc.GetChildNodes(NodeType.Comment, true);

            if (commentIndex < 0 || commentIndex >= comments.Count)
                throw new ArgumentException(
                    $"Comment index {commentIndex} is out of range (document has {comments.Count} comments)");

            var commentToDelete = comments[commentIndex] as Comment;
            if (commentToDelete == null)
                throw new InvalidOperationException($"Unable to find comment at index {commentIndex}");

            var author = commentToDelete.Author;
            var commentText = commentToDelete.GetText().Trim();
            var preview = commentText.Length > 50 ? commentText.Substring(0, 50) + "..." : commentText;

            var rangeStarts = doc.GetChildNodes(NodeType.CommentRangeStart, true).Cast<CommentRangeStart>();
            var rangeEnds = doc.GetChildNodes(NodeType.CommentRangeEnd, true).Cast<CommentRangeEnd>();

            foreach (var rangeStart in rangeStarts)
                if (rangeStart.Id == commentToDelete.Id)
                    rangeStart.Remove();

            foreach (var rangeEnd in rangeEnds)
                if (rangeEnd.Id == commentToDelete.Id)
                    rangeEnd.Remove();

            commentToDelete.Remove();

            doc.Save(outputPath);

            var remainingCount = doc.GetChildNodes(NodeType.Comment, true).Count;

            var result = $"Comment #{commentIndex} deleted successfully\n";
            result += $"Author: {author}\n";
            result += $"Content preview: {preview}\n";
            result += $"Remaining comments in document: {remainingCount}\n";
            result += $"Output: {outputPath}";

            return result;
        });
    }

    /// <summary>
    ///     Gets all top-level comments from the document
    /// </summary>
    /// <param name="_">Unused parameter</param>
    /// <param name="path">Word document file path</param>
    /// <returns>Formatted string with all comments and their replies</returns>
    private Task<string> GetCommentsAsync(JsonObject? _, string path)
    {
        return Task.Run(() =>
        {
            var doc = new Document(path);
            var allComments = doc.GetChildNodes(NodeType.Comment, true).Cast<Comment>().ToList();

            // Build set of reply comment IDs (comments that are replies to other comments)
            var replyCommentIds = new HashSet<int>();
            foreach (var comment in allComments)
                if (comment.Replies is { Count: > 0 })
                    foreach (var reply in comment.Replies.Cast<Comment>())
                        replyCommentIds.Add(reply.Id);

            // Filter to get only top-level comments (no ancestor and not in Replies collection)
            var topLevelComments = new List<Comment>();
            foreach (var comment in allComments)
                if (comment.Ancestor == null && !replyCommentIds.Contains(comment.Id))
                    if (topLevelComments.All(c => c.Id != comment.Id))
                        topLevelComments.Add(comment);

            // Also check comments from CommentRangeStart nodes
            var rangeStarts = doc.GetChildNodes(NodeType.CommentRangeStart, true).Cast<CommentRangeStart>();
            foreach (var rangeStart in rangeStarts)
            {
                var commentById = allComments.FirstOrDefault(c => c.Id == rangeStart.Id);
                if (commentById is { Ancestor: null } &&
                    !replyCommentIds.Contains(commentById.Id) &&
                    topLevelComments.All(c => c.Id != commentById.Id))
                    topLevelComments.Add(commentById);
            }

            if (topLevelComments.Count == 0) return "No comments found in document";

            var result = new StringBuilder();
            result.AppendLine($"Found {topLevelComments.Count} top-level comments:\n");

            var index = 0;
            foreach (var comment in topLevelComments)
            {
                AppendCommentInfo(result, comment, doc, index, 0);
                index++;
            }

            return result.ToString().TrimEnd();
        });
    }

    private void AppendCommentInfo(StringBuilder result, Comment comment, Document doc, int index, int indentLevel)
    {
        var indent = new string(' ', indentLevel * 2);
        var prefix = indentLevel == 0 ? $"Comment #{index}" : "  �|�w Reply";

        result.AppendLine($"{indent}{prefix}:");
        result.AppendLine($"{indent}  Author: {comment.Author}");
        result.AppendLine($"{indent}  Initial: {comment.Initial}");
        result.AppendLine($"{indent}  Date: {comment.DateTime:yyyy-MM-dd HH:mm:ss}");

        var commentText = comment.GetText().Trim();
        if (commentText.Length > 100) commentText = commentText.Substring(0, 100) + "...";
        result.AppendLine($"{indent}  Content: {commentText}");

        // Get commented text range if available
        var commentRangeStarts = doc.GetChildNodes(NodeType.CommentRangeStart, true).Cast<CommentRangeStart>();
        var hasRange = false;
        foreach (var rangeStart in commentRangeStarts)
            if (rangeStart.Id == comment.Id)
            {
                result.AppendLine($"{indent}  Range: Marked text");
                hasRange = true;
                break;
            }

        if (!hasRange) result.AppendLine($"{indent}  Range: Range marker not found");

        if (comment.Replies is { Count: > 0 })
        {
            result.AppendLine($"{indent}  Replies: {comment.Replies.Count}");
            foreach (var reply in comment.Replies.Cast<Comment>())
                AppendCommentInfo(result, reply, doc, -1, indentLevel + 1);
        }

        result.AppendLine();
    }

    /// <summary>
    ///     Adds a reply to an existing comment
    /// </summary>
    /// <param name="arguments">JSON arguments containing commentIndex, replyText or text, optional author, outputPath</param>
    /// <param name="path">Word document file path</param>
    /// <returns>Success message with reply details</returns>
    private Task<string> ReplyCommentAsync(JsonObject? arguments, string path)
    {
        return Task.Run(() =>
        {
            var outputPath = ArgumentHelper.GetStringNullable(arguments, "outputPath") ?? path;
            var commentIndex = ArgumentHelper.GetInt(arguments, "commentIndex");
            // Accept both replyText and text for compatibility
            var replyText = ArgumentHelper.GetString(arguments, "replyText", "text", "replyText or text");
            var author = ArgumentHelper.GetString(arguments, "author", "Reply Author");

            SecurityHelper.ValidateFilePath(outputPath, "outputPath", true);

            var doc = new Document(path);

            // Get all comments and filter to top-level comments only (same logic as GetCommentsAsync)
            var allComments = doc.GetChildNodes(NodeType.Comment, true).Cast<Comment>().ToList();

            // Build set of reply comment IDs
            var replyCommentIds = new HashSet<int>();
            foreach (var comment in allComments)
                if (comment.Replies is { Count: > 0 })
                    foreach (var reply in comment.Replies.Cast<Comment>())
                        replyCommentIds.Add(reply.Id);

            // Filter to get only top-level comments (no ancestor and not in Replies collection)
            var topLevelComments = new List<Comment>();
            foreach (var comment in allComments)
                if (comment.Ancestor == null && !replyCommentIds.Contains(comment.Id))
                    if (topLevelComments.All(c => c.Id != comment.Id))
                        topLevelComments.Add(comment);

            // Also check comments from CommentRangeStart nodes
            var rangeStarts = doc.GetChildNodes(NodeType.CommentRangeStart, true).Cast<CommentRangeStart>();
            foreach (var rangeStart in rangeStarts)
            {
                var commentById = allComments.FirstOrDefault(c => c.Id == rangeStart.Id);
                if (commentById is { Ancestor: null } &&
                    !replyCommentIds.Contains(commentById.Id) &&
                    topLevelComments.All(c => c.Id != commentById.Id))
                    topLevelComments.Add(commentById);
            }

            // Validate commentIndex against top-level comments only
            if (commentIndex < 0 || commentIndex >= topLevelComments.Count)
                throw new ArgumentException(
                    $"Comment index {commentIndex} is out of range (document has {topLevelComments.Count} top-level comments)");

            var parentComment = topLevelComments[commentIndex];

            // Check if parentComment is actually a reply (should not happen, but safety check)
            if (parentComment.Ancestor != null)
                throw new InvalidOperationException(
                    $"Comment index {commentIndex} points to a reply. Cannot add reply to a reply. Please use the top-level comment index.");

            // Use AddReply() method to create reply comment
            // It does NOT insert the reply content into the document body
            var initial = author.Length >= 2 ? author.Substring(0, 2).ToUpper() : author.ToUpper();
            _ = parentComment.AddReply(author, initial, DateTime.Now, replyText);
            doc.Save(outputPath);

            var result = $"Reply to comment #{commentIndex} added successfully\n";
            result += $"Original comment author: {parentComment.Author}\n";
            result += $"Reply author: {author}\n";
            result += $"Reply content: {replyText}\n";
            result += $"Output: {outputPath}";

            return result;
        });
    }
}