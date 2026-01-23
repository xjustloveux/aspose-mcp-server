using Aspose.Words;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Results.Common;

namespace AsposeMcpServer.Handlers.Word.Comment;

/// <summary>
///     Handler for deleting comments from Word documents.
/// </summary>
[ResultType(typeof(SuccessResult))]
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
    /// <exception cref="ArgumentException">Thrown when commentIndex is not provided or is out of range.</exception>
    /// <exception cref="InvalidOperationException">Thrown when comment cannot be found at the specified index.</exception>
    public override object Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var p = ExtractDeleteParameters(parameters);

        var doc = context.Document;
        var comments = doc.GetChildNodes(NodeType.Comment, true);

        if (p.CommentIndex < 0 || p.CommentIndex >= comments.Count)
            throw new ArgumentException(
                $"Comment index {p.CommentIndex} is out of range (document has {comments.Count} comments)");

        var commentToDelete = comments[p.CommentIndex] as Aspose.Words.Comment;
        if (commentToDelete == null)
            throw new InvalidOperationException($"Unable to find comment at index {p.CommentIndex}");

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

        return new SuccessResult
        {
            Message =
                $"Comment #{p.CommentIndex} deleted successfully\nAuthor: {author}\nRemaining comments: {remainingCount}"
        };
    }

    /// <summary>
    ///     Extracts and validates parameters for the delete comment operation.
    /// </summary>
    /// <param name="parameters">The operation parameters.</param>
    /// <returns>The extracted parameters.</returns>
    /// <exception cref="ArgumentException">Thrown when commentIndex is not provided.</exception>
    private static DeleteParameters ExtractDeleteParameters(OperationParameters parameters)
    {
        var commentIndex = parameters.GetOptional<int?>("commentIndex");

        if (!commentIndex.HasValue)
            throw new ArgumentException("commentIndex is required for delete operation");

        return new DeleteParameters(commentIndex.Value);
    }

    /// <summary>
    ///     Parameters for the delete comment operation.
    /// </summary>
    /// <param name="CommentIndex">The index of the comment to delete.</param>
    private sealed record DeleteParameters(int CommentIndex);
}
