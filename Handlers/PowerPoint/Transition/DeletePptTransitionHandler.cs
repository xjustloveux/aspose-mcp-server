using Aspose.Slides;
using Aspose.Slides.SlideShow;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Helpers;

namespace AsposeMcpServer.Handlers.PowerPoint.Transition;

/// <summary>
///     Handler for removing slide transitions from PowerPoint presentations.
/// </summary>
public class DeletePptTransitionHandler : OperationHandlerBase<Presentation>
{
    /// <inheritdoc />
    public override string Operation => "delete";

    /// <summary>
    ///     Removes the transition effect from a slide.
    /// </summary>
    /// <param name="context">The presentation context.</param>
    /// <param name="parameters">
    ///     Optional: slideIndex (default: 0).
    /// </param>
    /// <returns>Success message with deletion details.</returns>
    public override string Execute(OperationContext<Presentation> context, OperationParameters parameters)
    {
        var deleteParams = ExtractDeleteTransitionParameters(parameters);

        var presentation = context.Document;
        var slide = PowerPointHelper.GetSlide(presentation, deleteParams.SlideIndex);

        slide.SlideShowTransition.Type = TransitionType.None;
        slide.SlideShowTransition.AdvanceOnClick = true;
        slide.SlideShowTransition.AdvanceAfter = false;
        slide.SlideShowTransition.AdvanceAfterTime = 0;

        MarkModified(context);

        return Success($"Transition removed from slide {deleteParams.SlideIndex}.");
    }

    /// <summary>
    ///     Extracts delete transition parameters from operation parameters.
    /// </summary>
    /// <param name="parameters">The operation parameters.</param>
    /// <returns>The extracted delete transition parameters.</returns>
    private static DeleteTransitionParameters ExtractDeleteTransitionParameters(OperationParameters parameters)
    {
        return new DeleteTransitionParameters(
            parameters.GetOptional("slideIndex", 0)
        );
    }

    /// <summary>
    ///     Record for holding delete transition parameters.
    /// </summary>
    /// <param name="SlideIndex">The slide index.</param>
    private sealed record DeleteTransitionParameters(int SlideIndex);
}
