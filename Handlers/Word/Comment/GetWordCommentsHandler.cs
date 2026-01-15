using System.Text.Json;
using Aspose.Words;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Helpers;

namespace AsposeMcpServer.Handlers.Word.Comment;

/// <summary>
///     Handler for getting comments from Word documents.
/// </summary>
public class GetWordCommentsHandler : OperationHandlerBase<Document>
{
    /// <inheritdoc />
    public override string Operation => "get";

    /// <summary>
    ///     Gets all comments from the document as JSON with their replies.
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="parameters">No parameters required.</param>
    /// <returns>A JSON string containing all comments with their replies.</returns>
    public override string Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var doc = context.Document;
        var topLevelComments = WordCommentHelper.GetTopLevelComments(doc);

        if (topLevelComments.Count == 0)
            return JsonSerializer.Serialize(new
                { count = 0, comments = Array.Empty<object>(), message = "No comments found" });

        List<object> commentList = [];
        var index = 0;
        foreach (var comment in topLevelComments)
        {
            commentList.Add(WordCommentHelper.BuildCommentInfo(comment, doc, index));
            index++;
        }

        return JsonSerializer.Serialize(new { count = topLevelComments.Count, comments = commentList },
            JsonDefaults.Indented);
    }
}
