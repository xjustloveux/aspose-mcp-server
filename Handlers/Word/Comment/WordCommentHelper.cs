using Aspose.Words;

namespace AsposeMcpServer.Handlers.Word.Comment;

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
    public static List<Aspose.Words.Comment> GetTopLevelComments(Document doc)
    {
        var allComments = doc.GetChildNodes(NodeType.Comment, true).Cast<Aspose.Words.Comment>().ToList();
        var replyCommentIds = new HashSet<int>();

        foreach (var comment in allComments)
            if (comment.Replies is { Count: > 0 })
                foreach (var reply in comment.Replies.Cast<Aspose.Words.Comment>())
                    replyCommentIds.Add(reply.Id);

        List<Aspose.Words.Comment> topLevelComments = [];
        foreach (var comment in allComments)
            if (comment.Ancestor == null && !replyCommentIds.Contains(comment.Id))
                if (topLevelComments.All(c => c.Id != comment.Id))
                    topLevelComments.Add(comment);

        return topLevelComments.OrderBy(c => c.DateTime).ToList();
    }

    /// <summary>
    ///     Builds comment information object including nested replies.
    /// </summary>
    /// <param name="comment">The comment to build information for.</param>
    /// <param name="doc">The Word document.</param>
    /// <param name="index">The comment index.</param>
    /// <returns>An anonymous object containing the comment information.</returns>
    public static object BuildCommentInfo(Aspose.Words.Comment comment, Document doc, int index)
    {
        var commentText = comment.GetText().Trim();
        var rangeStarts = doc.GetChildNodes(NodeType.CommentRangeStart, true).Cast<CommentRangeStart>();
        var hasRange = rangeStarts.Any(rs => rs.Id == comment.Id);

        List<object> replies = [];
        if (comment.Replies is { Count: > 0 })
            foreach (var reply in comment.Replies.Cast<Aspose.Words.Comment>())
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
}
