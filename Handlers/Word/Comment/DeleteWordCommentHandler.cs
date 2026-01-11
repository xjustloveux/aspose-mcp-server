using Aspose.Words;
using AsposeMcpServer.Core.Handlers;

namespace AsposeMcpServer.Handlers.Word.Comment;

/// <summary>
///     Handler for deleting comments from Word documents.
/// </summary>
public class DeleteWordCommentHandler : OperationHandlerBase<Document>
{
    /// <inheritdoc />
    public override string Operation => "delete";

    /// <summary>
    ///     Deletes a comment from the document by index, including its range markers.
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="parameters">
    ///     Required: commentIndex
    /// </param>
    /// <returns>Success message with deletion details.</returns>
    public override string Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var commentIndex = parameters.GetOptional<int?>("commentIndex");

        if (!commentIndex.HasValue)
            throw new ArgumentException("commentIndex is required for delete operation");

        var doc = context.Document;
        var comments = doc.GetChildNodes(NodeType.Comment, true);

        if (commentIndex.Value < 0 || commentIndex.Value >= comments.Count)
            throw new ArgumentException(
                $"Comment index {commentIndex.Value} is out of range (document has {comments.Count} comments)");

        var commentToDelete = comments[commentIndex.Value] as Aspose.Words.Comment;
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

        MarkModified(context);

        var remainingCount = doc.GetChildNodes(NodeType.Comment, true).Count;

        return
            $"Comment #{commentIndex.Value} deleted successfully\nAuthor: {author}\nRemaining comments: {remainingCount}";
    }
}
