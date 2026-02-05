using Aspose.Slides;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Helpers.PowerPoint;
using AsposeMcpServer.Results.PowerPoint.Comment;

namespace AsposeMcpServer.Handlers.PowerPoint.Comment;

/// <summary>
///     Handler for getting comments from a PowerPoint slide.
/// </summary>
[ResultType(typeof(GetCommentsPptResult))]
public class GetPptCommentsHandler : OperationHandlerBase<Presentation>
{
    /// <inheritdoc />
    public override string Operation => "get";

    /// <summary>
    ///     Gets top-level comments from a slide.
    /// </summary>
    /// <param name="context">The presentation context.</param>
    /// <param name="parameters">
    ///     Optional: slideIndex (default: 0).
    /// </param>
    /// <returns>List of top-level comments on the specified slide.</returns>
    public override object Execute(OperationContext<Presentation> context, OperationParameters parameters)
    {
        var slideIndex = parameters.GetOptional("slideIndex", 0);

        var presentation = context.Document;
        var slide = PowerPointHelper.GetSlide(presentation, slideIndex);

        var comments = new List<PptCommentInfo>();
        var commentIndex = 0;

        foreach (var author in presentation.CommentAuthors)
        {
            var slideComments = slide.GetSlideComments(author);
            foreach (var comment in slideComments)
                if (comment.ParentComment == null)
                {
                    var replyCount = CountReplies(comment, slide);
                    comments.Add(BuildCommentInfo(commentIndex++, comment, replyCount));
                }
        }

        return new GetCommentsPptResult
        {
            Count = comments.Count,
            SlideIndex = slideIndex,
            Items = comments,
            Message = $"Found {comments.Count} comment(s) on slide {slideIndex}."
        };
    }

    /// <summary>
    ///     Counts the number of replies to a comment on a slide.
    /// </summary>
    /// <param name="parentComment">The parent comment.</param>
    /// <param name="slide">The slide.</param>
    /// <returns>The number of replies.</returns>
    private static int CountReplies(IComment parentComment, ISlide slide)
    {
        var count = 0;
        var presentation = slide.Presentation;
        foreach (var author in presentation.CommentAuthors)
        {
            var slideComments = slide.GetSlideComments(author);
            foreach (var comment in slideComments)
                if (comment.ParentComment == parentComment)
                    count++;
        }

        return count;
    }

    /// <summary>
    ///     Builds a comment info record from a comment.
    /// </summary>
    /// <param name="index">The comment index.</param>
    /// <param name="comment">The comment.</param>
    /// <param name="replyCount">The number of replies.</param>
    /// <returns>The comment info record.</returns>
    internal static PptCommentInfo BuildCommentInfo(int index, IComment comment, int replyCount)
    {
        return new PptCommentInfo
        {
            Index = index,
            Author = comment.Author?.Name ?? "Unknown",
            Text = comment.Text ?? string.Empty,
            X = comment.Position.X,
            Y = comment.Position.Y,
            CreatedTime = comment.CreatedTime.ToString("o"),
            ReplyCount = replyCount
        };
    }
}
