using Aspose.Slides;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Helpers.PowerPoint;
using AsposeMcpServer.Results.Common;

namespace AsposeMcpServer.Handlers.PowerPoint.Comment;

/// <summary>
///     Handler for replying to a comment on a PowerPoint slide.
/// </summary>
[ResultType(typeof(SuccessResult))]
public class ReplyPptCommentHandler : OperationHandlerBase<Presentation>
{
    /// <inheritdoc />
    public override string Operation => "reply";

    /// <summary>
    ///     Adds a reply to an existing comment.
    /// </summary>
    /// <param name="context">The presentation context.</param>
    /// <param name="parameters">
    ///     Required: commentIndex, text, author.
    ///     Optional: slideIndex (default: 0).
    /// </param>
    /// <returns>Success message with reply details.</returns>
    /// <exception cref="ArgumentException">Thrown when the comment index is out of range or required parameters are missing.</exception>
    public override object Execute(OperationContext<Presentation> context, OperationParameters parameters)
    {
        var commentIndex = parameters.GetRequired<int>("commentIndex");
        var text = parameters.GetRequired<string>("text");
        var authorName = parameters.GetRequired<string>("author");
        var slideIndex = parameters.GetOptional("slideIndex", 0);

        var presentation = context.Document;
        var slide = PowerPointHelper.GetSlide(presentation, slideIndex);

        var allComments = DeletePptCommentHandler.CollectTopLevelComments(presentation, slide);

        if (commentIndex < 0 || commentIndex >= allComments.Count)
            throw new ArgumentException(
                $"Comment index {commentIndex} is out of range (slide has {allComments.Count} comments).");

        var parentComment = allComments[commentIndex];
        var replyAuthor = AddPptCommentHandler.FindOrCreateAuthor(presentation, authorName);

        var reply = replyAuthor.Comments.AddComment(text, slide, parentComment.Position, DateTime.UtcNow);
        reply.ParentComment = parentComment;

        MarkModified(context);

        return new SuccessResult
        {
            Message = $"Reply added to comment {commentIndex} on slide {slideIndex} by '{authorName}'."
        };
    }
}
