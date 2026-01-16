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
        var p = ExtractReplyParameters(parameters);

        var doc = context.Document;
        var topLevelComments = WordCommentHelper.GetTopLevelComments(doc);

        if (p.CommentIndex < 0 || p.CommentIndex >= topLevelComments.Count)
            throw new ArgumentException(
                $"Comment index {p.CommentIndex} is out of range (document has {topLevelComments.Count} top-level comments)");

        var parentComment = topLevelComments[p.CommentIndex];
        var initial = p.AuthorInitial ?? (p.Author.Length >= 2 ? p.Author[..2].ToUpper() : p.Author.ToUpper());

        parentComment.AddReply(p.Author, initial, DateTime.UtcNow, p.ReplyText);

        MarkModified(context);

        return
            $"Reply added to comment #{p.CommentIndex}\nOriginal author: {parentComment.Author}\nReply author: {p.Author}\nReply: {p.ReplyText}";
    }

    /// <summary>
    ///     Extracts and validates parameters for the reply comment operation.
    /// </summary>
    /// <param name="parameters">The operation parameters.</param>
    /// <returns>The extracted parameters.</returns>
    /// <exception cref="ArgumentException">Thrown when required parameters are not provided.</exception>
    private static ReplyParameters ExtractReplyParameters(OperationParameters parameters)
    {
        var commentIndex = parameters.GetOptional<int?>("commentIndex");
        var replyText = parameters.GetOptional<string?>("replyText");
        var author = parameters.GetOptional("author", "Reply Author");
        var authorInitial = parameters.GetOptional<string?>("authorInitial");

        if (!commentIndex.HasValue)
            throw new ArgumentException("commentIndex is required for reply operation");
        if (string.IsNullOrEmpty(replyText))
            throw new ArgumentException("text or replyText is required for reply operation");

        return new ReplyParameters(commentIndex.Value, replyText, author, authorInitial);
    }

    /// <summary>
    ///     Parameters for the reply comment operation.
    /// </summary>
    /// <param name="CommentIndex">The index of the comment to reply to.</param>
    /// <param name="ReplyText">The text of the reply.</param>
    /// <param name="Author">The author of the reply.</param>
    /// <param name="AuthorInitial">The author's initials.</param>
    private record ReplyParameters(int CommentIndex, string ReplyText, string Author, string? AuthorInitial);
}
