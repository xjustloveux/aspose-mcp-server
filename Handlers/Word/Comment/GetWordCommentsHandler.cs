using Aspose.Words;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Helpers.Word;
using AsposeMcpServer.Results.Word.Comment;

namespace AsposeMcpServer.Handlers.Word.Comment;

/// <summary>
///     Handler for getting comments from Word documents.
/// </summary>
[ResultType(typeof(GetCommentsResult))]
public class GetWordCommentsHandler : OperationHandlerBase<Document>
{
    /// <inheritdoc />
    public override string Operation => "get";

    /// <summary>
    ///     Gets all comments from the document as JSON with their replies.
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="parameters">No parameters required.</param>
    /// <returns>A GetCommentsResult containing all comments with their replies.</returns>
    public override object Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        _ = parameters;

        var doc = context.Document;
        var topLevelComments = WordCommentHelper.GetTopLevelComments(doc);

        if (topLevelComments.Count == 0)
            return new GetCommentsResult
            {
                Count = 0,
                Comments = [],
                Message = "No comments found"
            };

        List<CommentInfo> commentList = [];
        var index = 0;
        foreach (var comment in topLevelComments)
        {
            commentList.Add(WordCommentHelper.BuildCommentInfo(comment, doc, index));
            index++;
        }

        return new GetCommentsResult
        {
            Count = topLevelComments.Count,
            Comments = commentList
        };
    }
}
