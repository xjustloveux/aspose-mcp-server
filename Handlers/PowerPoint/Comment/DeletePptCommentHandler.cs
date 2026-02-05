using Aspose.Slides;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Helpers.PowerPoint;
using AsposeMcpServer.Results.Common;

namespace AsposeMcpServer.Handlers.PowerPoint.Comment;

/// <summary>
///     Handler for deleting a comment from a PowerPoint slide.
/// </summary>
[ResultType(typeof(SuccessResult))]
public class DeletePptCommentHandler : OperationHandlerBase<Presentation>
{
    /// <inheritdoc />
    public override string Operation => "delete";

    /// <summary>
    ///     Deletes a comment from a slide.
    /// </summary>
    /// <param name="context">The presentation context.</param>
    /// <param name="parameters">
    ///     Required: commentIndex.
    ///     Optional: slideIndex (default: 0).
    /// </param>
    /// <returns>Success message with deletion details.</returns>
    /// <exception cref="ArgumentException">Thrown when the comment index is out of range.</exception>
    public override object Execute(OperationContext<Presentation> context, OperationParameters parameters)
    {
        var commentIndex = parameters.GetRequired<int>("commentIndex");
        var slideIndex = parameters.GetOptional("slideIndex", 0);

        var presentation = context.Document;
        var slide = PowerPointHelper.GetSlide(presentation, slideIndex);

        var allComments = CollectTopLevelComments(presentation, slide);

        if (commentIndex < 0 || commentIndex >= allComments.Count)
            throw new ArgumentException(
                $"Comment index {commentIndex} is out of range (slide has {allComments.Count} comments).");

        var targetComment = allComments[commentIndex];
        targetComment.Remove();

        MarkModified(context);

        return new SuccessResult
        {
            Message = $"Comment {commentIndex} deleted from slide {slideIndex}."
        };
    }

    /// <summary>
    ///     Collects all top-level comments on a slide across all authors.
    /// </summary>
    /// <param name="presentation">The presentation.</param>
    /// <param name="slide">The slide to collect comments from.</param>
    /// <returns>List of top-level comments on the slide.</returns>
    internal static List<IComment> CollectTopLevelComments(Presentation presentation, ISlide slide)
    {
        var allComments = new List<IComment>();
        foreach (var author in presentation.CommentAuthors)
        {
            var slideComments = slide.GetSlideComments(author);
            foreach (var comment in slideComments)
                if (comment.ParentComment == null)
                    allComments.Add(comment);
        }

        return allComments;
    }
}
