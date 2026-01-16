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
        var getParams = ExtractGetTransitionParameters(parameters);

        var presentation = context.Document;
        var slide = PowerPointHelper.GetSlide(presentation, getParams.SlideIndex);

        var transition = slide.SlideShowTransition;
        var hasTransition = transition.Type != TransitionType.None;

        var result = new
        {
            slideIndex = getParams.SlideIndex,
            type = transition.Type.ToString(),
            hasTransition,
            speed = transition.Speed.ToString(),
            advanceOnClick = transition.AdvanceOnClick,
            advanceAfter = transition.AdvanceAfter,
            advanceAfterSeconds = transition.AdvanceAfterTime / 1000.0
        };

        return JsonResult(result);
    }

    /// <summary>
    ///     Extracts get transition parameters from operation parameters.
    /// </summary>
    /// <param name="parameters">The operation parameters.</param>
    /// <returns>The extracted get transition parameters.</returns>
    private static GetTransitionParameters ExtractGetTransitionParameters(OperationParameters parameters)
    {
        return new GetTransitionParameters(
            parameters.GetOptional("slideIndex", 0)
        );
    }

    /// <summary>
    ///     Record for holding get transition parameters.
    /// </summary>
    /// <param name="SlideIndex">The slide index.</param>
    private record GetTransitionParameters(int SlideIndex);
}
