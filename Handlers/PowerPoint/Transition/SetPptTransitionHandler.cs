using Aspose.Slides;
using Aspose.Slides.SlideShow;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Helpers;

namespace AsposeMcpServer.Handlers.PowerPoint.Transition;

/// <summary>
///     Handler for setting slide transitions in PowerPoint presentations.
/// </summary>
public class SetPptTransitionHandler : OperationHandlerBase<Presentation>
{
    /// <summary>
    ///     Valid transition types for validation.
    /// </summary>
    private static readonly HashSet<string> ValidTransitionTypes = new(StringComparer.OrdinalIgnoreCase)
    {
        "none", "blinds", "checker", "circle", "comb", "cover", "cut", "diamond",
        "dissolve", "fade", "newsflash", "plus", "push", "random", "randombar",
        "split", "strips", "wedge", "wheel", "wipe", "zoom"
    };

    /// <inheritdoc />
    public override string Operation => "set";

    /// <summary>
    ///     Sets the transition effect for a slide.
    /// </summary>
    /// <param name="context">The presentation context.</param>
    /// <param name="parameters">
    ///     Required: transitionType.
    ///     Optional: slideIndex (default: 0), advanceAfterSeconds.
    /// </param>
    /// <returns>Success message with transition details.</returns>
    public override string Execute(OperationContext<Presentation> context, OperationParameters parameters)
    {
        var slideIndex = parameters.GetOptional("slideIndex", 0);
        var type = parameters.GetRequired<string>("transitionType");
        var advanceAfterSeconds = parameters.GetOptional<double?>("advanceAfterSeconds");

        if (string.IsNullOrWhiteSpace(type))
            throw new ArgumentException("transitionType is required");

        if (!ValidTransitionTypes.Contains(type))
            throw new ArgumentException(
                $"Invalid transition type: '{type}'. Valid types: {string.Join(", ", ValidTransitionTypes)}");

        var presentation = context.Document;
        var slide = PowerPointHelper.GetSlide(presentation, slideIndex);

        var transitionType = type.ToLower() switch
        {
            "none" => TransitionType.None,
            "blinds" => TransitionType.Blinds,
            "checker" => TransitionType.Checker,
            "circle" => TransitionType.Circle,
            "comb" => TransitionType.Comb,
            "cover" => TransitionType.Cover,
            "cut" => TransitionType.Cut,
            "diamond" => TransitionType.Diamond,
            "dissolve" => TransitionType.Dissolve,
            "fade" => TransitionType.Fade,
            "newsflash" => TransitionType.Newsflash,
            "plus" => TransitionType.Plus,
            "push" => TransitionType.Push,
            "random" => TransitionType.Random,
            "randombar" => TransitionType.RandomBar,
            "split" => TransitionType.Split,
            "strips" => TransitionType.Strips,
            "wedge" => TransitionType.Wedge,
            "wheel" => TransitionType.Wheel,
            "wipe" => TransitionType.Wipe,
            "zoom" => TransitionType.Zoom,
            _ => TransitionType.Fade
        };

        slide.SlideShowTransition.Type = transitionType;

        if (advanceAfterSeconds.HasValue)
        {
            var milliseconds = (uint)(advanceAfterSeconds.Value * 1000);
            slide.SlideShowTransition.AdvanceAfterTime = milliseconds;
            slide.SlideShowTransition.AdvanceAfter = milliseconds > 0;
        }

        MarkModified(context);

        return Success($"Transition '{type}' set for slide {slideIndex}.");
    }
}
