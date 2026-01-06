using System.ComponentModel;
using System.Text.Json;
using Aspose.Words;
using AsposeMcpServer.Core.Session;
using ModelContextProtocol.Server;

namespace AsposeMcpServer.Tools.Word;

/// <summary>
///     Unified tool for managing Word comments (add, delete, get, reply)
/// </summary>
[McpServerToolType]
public class WordCommentTool
{
    /// <summary>
    ///     Identity accessor for session isolation
    /// </summary>
    private readonly ISessionIdentityAccessor? _identityAccessor;

    /// <summary>
    ///     Session manager for document session operations
    /// </summary>
    private readonly DocumentSessionManager? _sessionManager;

    /// <summary>
    ///     Initializes a new instance of the WordCommentTool class
    /// </summary>
    /// <param name="sessionManager">Optional session manager for in-memory document operations</param>
    /// <param name="identityAccessor">Optional identity accessor for session isolation</param>
    public WordCommentTool(DocumentSessionManager? sessionManager = null,
        ISessionIdentityAccessor? identityAccessor = null)
    {
        _sessionManager = sessionManager;
        _identityAccessor = identityAccessor;
    }

    /// <summary>
    ///     Executes a Word comment operation (add, delete, get, reply).
    /// </summary>
    /// <param name="operation">The operation to perform: add, delete, get, reply.</param>
    /// <param name="path">Word document file path (required if no sessionId).</param>
    /// <param name="sessionId">Session ID for in-memory editing.</param>
    /// <param name="outputPath">Output file path (file mode only).</param>
    /// <param name="text">Comment text content.</param>
    /// <param name="author">Comment author name.</param>
    /// <param name="authorInitial">Author initials.</param>
    /// <param name="paragraphIndex">Paragraph index (0-based).</param>
    /// <param name="startRunIndex">Start run index.</param>
    /// <param name="endRunIndex">End run index.</param>
    /// <param name="commentIndex">Comment index (0-based).</param>
    /// <param name="replyText">Reply text content.</param>
    /// <returns>A message indicating the result of the operation, or JSON data for get operations.</returns>
    /// <exception cref="ArgumentException">Thrown when required parameters are missing or the operation is unknown.</exception>
    [McpServerTool(Name = "word_comment")]
    [Description(@"Manage Word comments. Supports 4 operations: add, delete, get, reply.

Usage examples:
- Add comment: word_comment(operation='add', path='doc.docx', text='This is a comment', author='Author Name')
- Delete comment: word_comment(operation='delete', path='doc.docx', commentIndex=0)
- Get all comments: word_comment(operation='get', path='doc.docx')
- Reply to comment: word_comment(operation='reply', path='doc.docx', commentIndex=0, text='This is a reply')")]
    public string Execute(
        [Description("Operation: add, delete, get, reply")]
        string operation,
        [Description("Document file path (required if no sessionId)")]
        string? path = null,
        [Description("Session ID for in-memory editing")]
        string? sessionId = null,
        [Description("Output file path (file mode only)")]
        string? outputPath = null,
        [Description("Comment text content")] string? text = null,
        [Description("Comment author name")] string? author = null,
        [Description("Author initials")] string? authorInitial = null,
        [Description("Paragraph index (0-based)")]
        int? paragraphIndex = null,
        [Description("Start run index")] int? startRunIndex = null,
        [Description("End run index")] int? endRunIndex = null,
        [Description("Comment index (0-based)")]
        int? commentIndex = null,
        [Description("Reply text content")] string? replyText = null)
    {
        using var ctx = DocumentContext<Document>.Create(_sessionManager, sessionId, path, _identityAccessor);

        return operation.ToLower() switch
        {
            "add" => AddComment(ctx, outputPath, text, author ?? "Comment Author", authorInitial, paragraphIndex,
                startRunIndex, endRunIndex),
            "delete" => DeleteComment(ctx, outputPath, commentIndex),
            "get" => GetComments(ctx),
            "reply" => ReplyComment(ctx, outputPath, commentIndex, replyText ?? text, author ?? "Reply Author",
                authorInitial),
            _ => throw new ArgumentException($"Unknown operation: {operation}")
        };
    }

    /// <summary>
    ///     Adds a comment to the document at the specified paragraph and run range.
    /// </summary>
    /// <param name="ctx">The document context.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="text">The comment text content.</param>
    /// <param name="author">The comment author name.</param>
    /// <param name="authorInitial">The author initials.</param>
    /// <param name="paragraphIndex">The paragraph index (0-based, -1 for last paragraph).</param>
    /// <param name="startRunIndex">The start run index within the paragraph.</param>
    /// <param name="endRunIndex">The end run index within the paragraph.</param>
    /// <returns>A message indicating the result of the operation.</returns>
    /// <exception cref="ArgumentException">Thrown when text is empty or run index is invalid.</exception>
    /// <exception cref="InvalidOperationException">Thrown when unable to determine target paragraph or comment range.</exception>
    private static string AddComment(DocumentContext<Document> ctx, string? outputPath, string? text, string author,
        string? authorInitial, int? paragraphIndex, int? startRunIndex, int? endRunIndex)
    {
        if (string.IsNullOrEmpty(text))
            throw new ArgumentException("text is required for add operation");

        var doc = ctx.Document;
        List<Paragraph> paragraphs = [];
        foreach (var section in doc.Sections.Cast<Section>())
        {
            var bodyParagraphs = section.Body.GetChildNodes(NodeType.Paragraph, false).Cast<Paragraph>().ToList();
            paragraphs.AddRange(bodyParagraphs);
        }

        Paragraph? targetPara;
        Run? startRun;
        Run? endRun;

        if (paragraphIndex.HasValue)
        {
            if (paragraphIndex.Value == -1)
            {
                if (paragraphs.Count > 0)
                    targetPara = paragraphs[^1];
                else
                    throw new InvalidOperationException("Document has no paragraphs to add comment to");
            }
            else if (paragraphIndex.Value < 0 || paragraphIndex.Value >= paragraphs.Count)
            {
                throw new ArgumentException(
                    $"Paragraph index {paragraphIndex.Value} is out of range (document has {paragraphs.Count} paragraphs)");
            }
            else
            {
                targetPara = paragraphs[paragraphIndex.Value];
            }
        }
        else
        {
            var builder = new DocumentBuilder(doc);
            builder.MoveToDocumentEnd();

            var newPara = new Paragraph(doc);
            var newRun = new Run(doc, " ");
            newPara.AppendChild(newRun);
            doc.LastSection.Body.AppendChild(newPara);

            targetPara = newPara;
        }

        if (targetPara == null)
            throw new InvalidOperationException("Unable to determine target paragraph");

        var runs = targetPara.GetChildNodes(NodeType.Run, false);
        if (runs == null || runs.Count == 0)
        {
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
        if (para == null)
            throw new InvalidOperationException("Unable to find paragraph node containing Run");

        var initial = authorInitial ?? (author.Length >= 2 ? author[..2].ToUpper() : author.ToUpper());
        var comment = new Comment(doc, author, initial, DateTime.UtcNow);
        comment.Paragraphs.Add(new Paragraph(doc));
        comment.FirstParagraph.AppendChild(new Run(doc, text));

        var rangeStart = new CommentRangeStart(doc, comment.Id);
        var rangeEnd = new CommentRangeEnd(doc, comment.Id);

        var startPara = para;
        if (startRun.ParentNode != startPara)
            if (startRun.ParentNode is Paragraph parentPara)
                startPara = parentPara;

        startPara.InsertBefore(rangeStart, startRun.ParentNode == startPara ? startRun : startPara.FirstChild);

        var endPara = endRun.ParentNode as Paragraph ?? endRun.GetAncestor(NodeType.Paragraph) as Paragraph;
        if (endPara == null)
            throw new InvalidOperationException("Unable to find paragraph containing endRun");

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

        doc.EnsureMinimum();
        ctx.Save(outputPath);

        var result = "Comment added successfully\n";
        result += $"Author: {author}\n";
        result += $"Content: {text}\n";
        result += ctx.GetOutputMessage(outputPath);

        return result;
    }

    /// <summary>
    ///     Deletes a comment from the document by index, including its range markers.
    /// </summary>
    /// <param name="ctx">The document context.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="commentIndex">The comment index (0-based).</param>
    /// <returns>A message indicating the result of the operation.</returns>
    /// <exception cref="ArgumentException">Thrown when comment index is not provided or out of range.</exception>
    /// <exception cref="InvalidOperationException">Thrown when unable to find the comment.</exception>
    private static string DeleteComment(DocumentContext<Document> ctx, string? outputPath, int? commentIndex)
    {
        if (!commentIndex.HasValue)
            throw new ArgumentException("commentIndex is required for delete operation");

        var doc = ctx.Document;
        var comments = doc.GetChildNodes(NodeType.Comment, true);

        if (commentIndex.Value < 0 || commentIndex.Value >= comments.Count)
            throw new ArgumentException(
                $"Comment index {commentIndex.Value} is out of range (document has {comments.Count} comments)");

        var commentToDelete = comments[commentIndex.Value] as Comment;
        if (commentToDelete == null)
            throw new InvalidOperationException($"Unable to find comment at index {commentIndex.Value}");

        var author = commentToDelete.Author;

        var rangeStarts = doc.GetChildNodes(NodeType.CommentRangeStart, true).Cast<CommentRangeStart>();
        var rangeEnds = doc.GetChildNodes(NodeType.CommentRangeEnd, true).Cast<CommentRangeEnd>();

        foreach (var rangeStart in rangeStarts)
            if (rangeStart.Id == commentToDelete.Id)
                rangeStart.Remove();

        foreach (var rangeEnd in rangeEnds)
            if (rangeEnd.Id == commentToDelete.Id)
                rangeEnd.Remove();

        commentToDelete.Remove();
        ctx.Save(outputPath);

        var remainingCount = doc.GetChildNodes(NodeType.Comment, true).Count;

        return
            $"Comment #{commentIndex.Value} deleted successfully\nAuthor: {author}\nRemaining comments: {remainingCount}\n{ctx.GetOutputMessage(outputPath)}";
    }

    /// <summary>
    ///     Gets all top-level comments from the document, excluding reply comments.
    /// </summary>
    /// <param name="doc">The Word document.</param>
    /// <returns>A list of top-level comments ordered by date.</returns>
    private static List<Comment> GetTopLevelComments(Document doc)
    {
        var allComments = doc.GetChildNodes(NodeType.Comment, true).Cast<Comment>().ToList();
        var replyCommentIds = new HashSet<int>();

        foreach (var comment in allComments)
            if (comment.Replies is { Count: > 0 })
                foreach (var reply in comment.Replies.Cast<Comment>())
                    replyCommentIds.Add(reply.Id);

        List<Comment> topLevelComments = [];
        foreach (var comment in allComments)
            if (comment.Ancestor == null && !replyCommentIds.Contains(comment.Id))
                if (topLevelComments.All(c => c.Id != comment.Id))
                    topLevelComments.Add(comment);

        return topLevelComments.OrderBy(c => c.DateTime).ToList();
    }

    /// <summary>
    ///     Gets all comments from the document as JSON with their replies.
    /// </summary>
    /// <param name="ctx">The document context.</param>
    /// <returns>A JSON string containing all comments with their replies.</returns>
    private static string GetComments(DocumentContext<Document> ctx)
    {
        var doc = ctx.Document;
        var topLevelComments = GetTopLevelComments(doc);

        if (topLevelComments.Count == 0)
            return JsonSerializer.Serialize(new
                { count = 0, comments = Array.Empty<object>(), message = "No comments found" });

        List<object> commentList = [];
        var index = 0;
        foreach (var comment in topLevelComments)
        {
            commentList.Add(BuildCommentInfo(comment, doc, index));
            index++;
        }

        return JsonSerializer.Serialize(new { count = topLevelComments.Count, comments = commentList },
            new JsonSerializerOptions { WriteIndented = true });
    }

    /// <summary>
    ///     Builds comment information object including nested replies.
    /// </summary>
    /// <param name="comment">The comment to build information for.</param>
    /// <param name="doc">The Word document.</param>
    /// <param name="index">The comment index.</param>
    /// <returns>An anonymous object containing the comment information.</returns>
    private static object BuildCommentInfo(Comment comment, Document doc, int index)
    {
        var commentText = comment.GetText().Trim();
        var rangeStarts = doc.GetChildNodes(NodeType.CommentRangeStart, true).Cast<CommentRangeStart>();
        var hasRange = rangeStarts.Any(rs => rs.Id == comment.Id);

        List<object> replies = [];
        if (comment.Replies is { Count: > 0 })
            foreach (var reply in comment.Replies.Cast<Comment>())
                replies.Add(BuildCommentInfo(reply, doc, -1));

        return new
        {
            index,
            author = comment.Author,
            initial = comment.Initial,
            date = comment.DateTime.ToString("yyyy-MM-dd HH:mm:ss"),
            content = commentText,
            hasRange,
            replyCount = comment.Replies?.Count ?? 0,
            replies
        };
    }

    /// <summary>
    ///     Adds a reply to an existing comment.
    /// </summary>
    /// <param name="ctx">The document context.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="commentIndex">The comment index to reply to (0-based).</param>
    /// <param name="replyText">The reply text content.</param>
    /// <param name="author">The reply author name.</param>
    /// <param name="authorInitial">The author initials.</param>
    /// <returns>A message indicating the result of the operation.</returns>
    /// <exception cref="ArgumentException">Thrown when comment index is not provided, out of range, or reply text is empty.</exception>
    private static string ReplyComment(DocumentContext<Document> ctx, string? outputPath, int? commentIndex,
        string? replyText, string author, string? authorInitial)
    {
        if (!commentIndex.HasValue)
            throw new ArgumentException("commentIndex is required for reply operation");
        if (string.IsNullOrEmpty(replyText))
            throw new ArgumentException("text or replyText is required for reply operation");

        var doc = ctx.Document;
        var topLevelComments = GetTopLevelComments(doc);

        if (commentIndex.Value < 0 || commentIndex.Value >= topLevelComments.Count)
            throw new ArgumentException(
                $"Comment index {commentIndex.Value} is out of range (document has {topLevelComments.Count} top-level comments)");

        var parentComment = topLevelComments[commentIndex.Value];
        var initial = authorInitial ?? (author.Length >= 2 ? author[..2].ToUpper() : author.ToUpper());

        parentComment.AddReply(author, initial, DateTime.UtcNow, replyText);
        ctx.Save(outputPath);

        return
            $"Reply added to comment #{commentIndex.Value}\nOriginal author: {parentComment.Author}\nReply author: {author}\nReply: {replyText}\n{ctx.GetOutputMessage(outputPath)}";
    }
}