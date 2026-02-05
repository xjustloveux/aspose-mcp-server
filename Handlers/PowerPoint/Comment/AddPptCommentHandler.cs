using System.Drawing;
using Aspose.Slides;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Helpers.PowerPoint;
using AsposeMcpServer.Results.Common;

namespace AsposeMcpServer.Handlers.PowerPoint.Comment;

/// <summary>
///     Handler for adding a comment to a PowerPoint slide.
/// </summary>
[ResultType(typeof(SuccessResult))]
public class AddPptCommentHandler : OperationHandlerBase<Presentation>
{
    /// <inheritdoc />
    public override string Operation => "add";

    /// <summary>
    ///     Adds a comment to a slide.
    /// </summary>
    /// <param name="context">The presentation context.</param>
    /// <param name="parameters">
    ///     Required: text, author.
    ///     Optional: slideIndex (default: 0), x (default: 0), y (default: 0).
    /// </param>
    /// <returns>Success message with comment addition details.</returns>
    /// <exception cref="ArgumentException">Thrown when required parameters are missing.</exception>
    public override object Execute(OperationContext<Presentation> context, OperationParameters parameters)
    {
        var text = parameters.GetRequired<string>("text");
        var authorName = parameters.GetRequired<string>("author");
        var slideIndex = parameters.GetOptional("slideIndex", 0);
        var x = parameters.GetOptional("x", 0f);
        var y = parameters.GetOptional("y", 0f);

        var presentation = context.Document;
        var slide = PowerPointHelper.GetSlide(presentation, slideIndex);

        var author = FindOrCreateAuthor(presentation, authorName);
        author.Comments.AddComment(text, slide, new PointF(x, y), DateTime.UtcNow);

        MarkModified(context);

        return new SuccessResult { Message = $"Comment added to slide {slideIndex} by '{authorName}'." };
    }

    /// <summary>
    ///     Finds an existing author or creates a new one.
    /// </summary>
    /// <param name="presentation">The presentation.</param>
    /// <param name="authorName">The author name.</param>
    /// <returns>The comment author.</returns>
    internal static ICommentAuthor FindOrCreateAuthor(Presentation presentation, string authorName)
    {
        foreach (var existingAuthor in presentation.CommentAuthors)
            if (existingAuthor.Name == authorName)
                return existingAuthor;

        var initials = string.Join("", authorName.Split(' ', StringSplitOptions.RemoveEmptyEntries)
            .Select(w => w[0].ToString().ToUpperInvariant()));
        if (string.IsNullOrEmpty(initials)) initials = "U";
        return presentation.CommentAuthors.AddAuthor(authorName, initials);
    }
}
