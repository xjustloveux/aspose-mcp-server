using Aspose.Words;
using AsposeMcpServer.Results.Word.Comment;

namespace AsposeMcpServer.Helpers.Word;

/// <summary>
///     Helper methods for Word comment operations.
/// </summary>
public static class WordCommentHelper
{
    /// <summary>
    ///     Gets all top-level comments from the document, excluding reply comments.
    /// </summary>
    /// <param name="doc">The Word document.</param>
    /// <returns>A list of top-level comments ordered by date.</returns>
    public static List<Comment> GetTopLevelComments(Document doc)
    {
        var allComments = doc.GetChildNodes(NodeType.Comment, true).Cast<Comment>().ToList();
        var replyCommentIds = allComments
            .Where(comment => comment.Replies is { Count: > 0 })
            .SelectMany(comment => comment.Replies.Cast<Comment>())
            .Select(reply => reply.Id)
            .ToHashSet();

        List<Comment> topLevelComments = [];
        foreach (var comment in allComments)
            if (comment.Ancestor == null && !replyCommentIds.Contains(comment.Id) &&
                topLevelComments.All(c => c.Id != comment.Id))
                topLevelComments.Add(comment);

        return topLevelComments.OrderBy(c => c.DateTime).ToList();
    }

    /// <summary>
    ///     Builds comment information object including nested replies.
    /// </summary>
    /// <param name="comment">The comment to build information for.</param>
    /// <param name="doc">The Word document.</param>
    /// <param name="index">The comment index.</param>
    /// <returns>A CommentInfo object containing the comment information.</returns>
    public static CommentInfo BuildCommentInfo(Comment comment, Document doc, int index)
    {
        var commentText = comment.GetText().Trim();
        var rangeStarts = doc.GetChildNodes(NodeType.CommentRangeStart, true).Cast<CommentRangeStart>();
        var hasRange = rangeStarts.Any(rs => rs.Id == comment.Id);

        List<CommentInfo> replies = [];
        if (comment.Replies is { Count: > 0 })
            foreach (var reply in comment.Replies.Cast<Comment>())
                replies.Add(BuildCommentInfo(reply, doc, -1));

        return new CommentInfo
        {
            Index = index,
            Author = comment.Author,
            Initial = comment.Initial,
            Date = comment.DateTime.ToString("yyyy-MM-dd HH:mm:ss"),
            Content = commentText,
            HasRange = hasRange,
            ReplyCount = comment.Replies?.Count ?? 0,
            Replies = replies
        };
    }
}
