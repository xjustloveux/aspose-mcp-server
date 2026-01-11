using Aspose.Words;
using AsposeMcpServer.Core.Handlers;

namespace AsposeMcpServer.Handlers.Word.Comment;

/// <summary>
///     Handler for replying to comments in Word documents.
/// </summary>
public class ReplyWordCommentHandler : OperationHandlerBase<Document>
{
    /// <inheritdoc />
    public override string Operation => "reply";

    /// <summary>
    ///     Adds a reply to an existing comment.
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="parameters">
    ///     Required: commentIndex, replyText
    ///     Optional: author, authorInitial
    /// </param>
    /// <returns>Success message with reply details.</returns>
    public override string Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var commentIndex = parameters.GetOptional<int?>("commentIndex");
        var replyText = parameters.GetOptional<string?>("replyText");
        var author = parameters.GetOptional("author", "Reply Author");
        var authorInitial = parameters.GetOptional<string?>("authorInitial");

        if (!commentIndex.HasValue)
            throw new ArgumentException("commentIndex is required for reply operation");
        if (string.IsNullOrEmpty(replyText))
            throw new ArgumentException("text or replyText is required for reply operation");

        var doc = context.Document;
        var topLevelComments = WordCommentHelper.GetTopLevelComments(doc);

        if (commentIndex.Value < 0 || commentIndex.Value >= topLevelComments.Count)
            throw new ArgumentException(
                $"Comment index {commentIndex.Value} is out of range (document has {topLevelComments.Count} top-level comments)");

        var parentComment = topLevelComments[commentIndex.Value];
        var initial = authorInitial ?? (author.Length >= 2 ? author[..2].ToUpper() : author.ToUpper());

        parentComment.AddReply(author, initial, DateTime.UtcNow, replyText);

        MarkModified(context);

        return
            $"Reply added to comment #{commentIndex.Value}\nOriginal author: {parentComment.Author}\nReply author: {author}\nReply: {replyText}";
    }
}
