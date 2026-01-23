using Aspose.Slides;
using Aspose.Slides.SlideShow;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Helpers.PowerPoint;
using AsposeMcpServer.Results.PowerPoint.Transition;

namespace AsposeMcpServer.Handlers.PowerPoint.Transition;

/// <summary>
///     Handler for getting slide transition information from PowerPoint presentations.
/// </summary>
[ResultType(typeof(GetTransitionResult))]
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
    public override object Execute(OperationContext<Presentation> context, OperationParameters parameters)
    {
        var getParams = ExtractGetTransitionParameters(parameters);

        var presentation = context.Document;
        var slide = PowerPointHelper.GetSlide(presentation, getParams.SlideIndex);

        var transition = slide.SlideShowTransition;
        var hasTransition = transition.Type != TransitionType.None;

        var result = new GetTransitionResult
        {
            SlideIndex = getParams.SlideIndex,
            Type = transition.Type.ToString(),
            HasTransition = hasTransition,
            Speed = transition.Speed.ToString(),
            AdvanceOnClick = transition.AdvanceOnClick,
            AdvanceAfter = transition.AdvanceAfter,
            AdvanceAfterSeconds = transition.AdvanceAfterTime / 1000.0
        };

        return result;
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
    private sealed record GetTransitionParameters(int SlideIndex);
}
