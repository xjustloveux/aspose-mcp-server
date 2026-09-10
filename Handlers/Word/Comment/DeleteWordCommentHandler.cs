using Aspose.Words;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Helpers.Word;
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

        // The index space must be the one get and reply use: top-level comments ordered by date.
        // A raw NodeType.Comment walk also returns replies and follows document order, so on a
        // document with replies or out-of-order dates the same index named a different comment
        // depending on which operation the caller had used, and delete removed the wrong one.
        var comments = WordCommentHelper.GetTopLevelComments(doc);

        if (p.CommentIndex < 0 || p.CommentIndex >= comments.Count)
            throw new ArgumentException(
                $"Comment index {p.CommentIndex} is out of range (document has {comments.Count} comments)");

        var commentToDelete = comments[p.CommentIndex];

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

        // Reported in the same space as the index the caller supplied: top-level comments only.
        var remainingCount = WordCommentHelper.GetTopLevelComments(doc).Count;

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
