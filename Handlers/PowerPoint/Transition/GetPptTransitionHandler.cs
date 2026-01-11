using Aspose.Slides;
using Aspose.Slides.SlideShow;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Helpers;

namespace AsposeMcpServer.Handlers.PowerPoint.Transition;

/// <summary>
///     Handler for getting slide transition information from PowerPoint presentations.
/// </summary>
public class GetPptTransitionHandler : OperationHandlerBase<Presentation>
{
    /// <inheritdoc />
    public override string Operation => "get";

    /// <summary>
    ///     Gets the transition information for a slide.
    /// </summary>
    /// <param name="context">The presentation context.</param>
    /// <param name="parameters">
    ///     Optional: slideIndex (default: 0).
    /// </param>
    /// <returns>JSON result with transition information.</returns>
    public override string Execute(OperationContext<Presentation> context, OperationParameters parameters)
    {
        var slideIndex = parameters.GetOptional("slideIndex", 0);

        var presentation = context.Document;
        var slide = PowerPointHelper.GetSlide(presentation, slideIndex);

        var transition = slide.SlideShowTransition;
        var hasTransition = transition.Type != TransitionType.None;

        var result = new
        {
            slideIndex,
            type = transition.Type.ToString(),
            hasTransition,
            speed = transition.Speed.ToString(),
            advanceOnClick = transition.AdvanceOnClick,
            advanceAfter = transition.AdvanceAfter,
            advanceAfterSeconds = transition.AdvanceAfterTime / 1000.0
        };

        return JsonResult(result);
    }
}
