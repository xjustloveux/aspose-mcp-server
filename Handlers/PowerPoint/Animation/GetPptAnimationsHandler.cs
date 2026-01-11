using Aspose.Slides;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Helpers;

namespace AsposeMcpServer.Handlers.PowerPoint.Animation;

/// <summary>
///     Handler for getting animation information from PowerPoint presentations.
/// </summary>
public class GetPptAnimationsHandler : OperationHandlerBase<Presentation>
{
    /// <inheritdoc />
    public override string Operation => "get";

    /// <summary>
    ///     Gets animation information for a slide or specific shape.
    /// </summary>
    /// <param name="context">The presentation context.</param>
    /// <param name="parameters">
    ///     Required: slideIndex
    ///     Optional: shapeIndex
    /// </param>
    /// <returns>JSON string containing the animation information.</returns>
    public override string Execute(OperationContext<Presentation> context, OperationParameters parameters)
    {
        var slideIndex = parameters.GetRequired<int>("slideIndex");
        var shapeIndex = parameters.GetOptional<int?>("shapeIndex");

        var presentation = context.Document;
        var slide = PowerPointHelper.GetSlide(presentation, slideIndex);
        var sequence = slide.Timeline.MainSequence;

        List<object> animations = [];
        var index = 0;

        foreach (var effect in sequence)
        {
            if (shapeIndex.HasValue)
            {
                var targetShapeIndex = slide.Shapes.IndexOf(effect.TargetShape);
                if (targetShapeIndex != shapeIndex.Value)
                    continue;
            }

            var shapeName = effect.TargetShape?.Name ?? "(unknown)";
            var shapeIdx = effect.TargetShape != null ? slide.Shapes.IndexOf(effect.TargetShape) : -1;

            animations.Add(new
            {
                index,
                shapeIndex = shapeIdx,
                shapeName,
                effectType = effect.Type.ToString(),
                effectSubtype = effect.Subtype.ToString(),
                triggerType = effect.Timing.TriggerType.ToString(),
                duration = effect.Timing.Duration,
                delay = effect.Timing.TriggerDelayTime
            });

            index++;
        }

        var result = new
        {
            slideIndex,
            filterByShapeIndex = shapeIndex,
            totalAnimationsOnSlide = sequence.Count,
            animations
        };

        return JsonResult(result);
    }
}
