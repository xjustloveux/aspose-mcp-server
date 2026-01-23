using Aspose.Slides;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Helpers.PowerPoint;
using AsposeMcpServer.Results.PowerPoint.Animation;

namespace AsposeMcpServer.Handlers.PowerPoint.Animation;

/// <summary>
///     Handler for getting animation information from PowerPoint presentations.
/// </summary>
[ResultType(typeof(GetAnimationsResult))]
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
    public override object Execute(OperationContext<Presentation> context, OperationParameters parameters)
    {
        var p = ExtractGetAnimationsParameters(parameters);

        var presentation = context.Document;
        var slide = PowerPointHelper.GetSlide(presentation, p.SlideIndex);
        var sequence = slide.Timeline.MainSequence;

        List<AnimationInfo> animations = [];
        var index = 0;

        foreach (var effect in sequence)
        {
            if (p.ShapeIndex.HasValue)
            {
                var targetShapeIndex = slide.Shapes.IndexOf(effect.TargetShape);
                if (targetShapeIndex != p.ShapeIndex.Value)
                    continue;
            }

            var shapeName = effect.TargetShape?.Name ?? "(unknown)";
            var shapeIdx = effect.TargetShape != null ? slide.Shapes.IndexOf(effect.TargetShape) : -1;

            animations.Add(new AnimationInfo
            {
                Index = index,
                ShapeIndex = shapeIdx,
                ShapeName = shapeName,
                EffectType = effect.Type.ToString(),
                EffectSubtype = effect.Subtype.ToString(),
                TriggerType = effect.Timing.TriggerType.ToString(),
                Duration = effect.Timing.Duration,
                Delay = effect.Timing.TriggerDelayTime
            });

            index++;
        }

        var result = new GetAnimationsResult
        {
            SlideIndex = p.SlideIndex,
            FilterByShapeIndex = p.ShapeIndex,
            TotalAnimationsOnSlide = sequence.Count,
            Animations = animations
        };

        return result;
    }

    /// <summary>
    ///     Extracts get animations parameters from operation parameters.
    /// </summary>
    /// <param name="parameters">The operation parameters.</param>
    /// <returns>The extracted get animations parameters.</returns>
    private static GetAnimationsParameters ExtractGetAnimationsParameters(OperationParameters parameters)
    {
        return new GetAnimationsParameters(
            parameters.GetRequired<int>("slideIndex"),
            parameters.GetOptional<int?>("shapeIndex"));
    }

    /// <summary>
    ///     Record for holding get animations parameters.
    /// </summary>
    /// <param name="SlideIndex">The slide index.</param>
    /// <param name="ShapeIndex">The optional shape index filter.</param>
    private sealed record GetAnimationsParameters(int SlideIndex, int? ShapeIndex);
}
