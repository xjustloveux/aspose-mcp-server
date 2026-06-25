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
        var perShapeIndex = new Dictionary<IShape, int>(ReferenceEqualityComparer.Instance);

        foreach (var effect in sequence)
        {
            var targetShape = effect.TargetShape;
            var shapeIdx = targetShape != null ? slide.Shapes.IndexOf(targetShape) : -1;

            // Index is the animation's position WITHIN its own shape — the same per-shape ordinal that
            // edit/delete consume as animationIndex (sequence.Where(e => e.TargetShape == shape)). It is
            // computed before the shape filter so it stays correct whether or not shapeIndex is given.
            var animIndex = 0;
            if (targetShape != null)
            {
                perShapeIndex.TryGetValue(targetShape, out animIndex);
                perShapeIndex[targetShape] = animIndex + 1;
            }

            if (p.ShapeIndex.HasValue && shapeIdx != p.ShapeIndex.Value)
                continue;

            animations.Add(new AnimationInfo
            {
                Index = animIndex,
                ShapeIndex = shapeIdx,
                ShapeName = targetShape?.Name ?? "(unknown)",
                EffectType = effect.Type.ToString(),
                EffectSubtype = effect.Subtype.ToString(),
                TriggerType = effect.Timing.TriggerType.ToString(),
                Duration = SanitizeFloat(effect.Timing.Duration),
                Delay = SanitizeFloat(effect.Timing.TriggerDelayTime)
            });
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
    ///     Sanitizes a float value to ensure it can be serialized to JSON.
    ///     Converts NaN and Infinity values to 0.
    /// </summary>
    /// <param name="value">The float value to sanitize.</param>
    /// <returns>The sanitized float value.</returns>
    private static float SanitizeFloat(float value)
    {
        return float.IsNaN(value) || float.IsInfinity(value) ? 0f : value;
    }

    /// <summary>
    ///     Record for holding get animations parameters.
    /// </summary>
    /// <param name="SlideIndex">The slide index.</param>
    /// <param name="ShapeIndex">The optional shape index filter.</param>
    private sealed record GetAnimationsParameters(int SlideIndex, int? ShapeIndex);
}
