using Aspose.Slides;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Helpers;

namespace AsposeMcpServer.Handlers.PowerPoint.Animation;

/// <summary>
///     Handler for editing animations in PowerPoint presentations.
/// </summary>
public class EditPptAnimationHandler : OperationHandlerBase<Presentation>
{
    /// <inheritdoc />
    public override string Operation => "edit";

    /// <summary>
    ///     Edits animation properties. If animationIndex is provided, modifies that specific animation;
    ///     otherwise replaces all animations for the shape.
    /// </summary>
    /// <param name="context">The presentation context.</param>
    /// <param name="parameters">
    ///     Required: slideIndex, shapeIndex
    ///     Optional: animationIndex, effectType, effectSubtype, triggerType, duration, delay
    /// </param>
    /// <returns>Success message with update details.</returns>
    public override string Execute(OperationContext<Presentation> context, OperationParameters parameters)
    {
        var slideIndex = parameters.GetRequired<int>("slideIndex");
        var shapeIndex = parameters.GetRequired<int>("shapeIndex");
        var animationIndex = parameters.GetOptional<int?>("animationIndex");
        var effectTypeStr = parameters.GetOptional<string?>("effectType");
        var effectSubtypeStr = parameters.GetOptional<string?>("effectSubtype");
        var triggerTypeStr = parameters.GetOptional<string?>("triggerType");
        var duration = parameters.GetOptional<float?>("duration");
        var delay = parameters.GetOptional<float?>("delay");

        var presentation = context.Document;
        var slide = PowerPointHelper.GetSlide(presentation, slideIndex);
        PowerPointHelper.ValidateShapeIndex(shapeIndex, slide);

        var shape = slide.Shapes[shapeIndex];
        var sequence = slide.Timeline.MainSequence;
        var animations = sequence.Where(e => e.TargetShape == shape).ToList();

        if (animationIndex.HasValue)
        {
            if (animationIndex.Value < 0 || animationIndex.Value >= animations.Count)
                throw new ArgumentException($"animationIndex must be between 0 and {animations.Count - 1}");

            var effect = animations[animationIndex.Value];
            if (duration.HasValue) effect.Timing.Duration = duration.Value;
            if (delay.HasValue) effect.Timing.TriggerDelayTime = delay.Value;

            if (!string.IsNullOrEmpty(effectTypeStr) || !string.IsNullOrEmpty(effectSubtypeStr) ||
                !string.IsNullOrEmpty(triggerTypeStr))
            {
                var newEffectType = !string.IsNullOrEmpty(effectTypeStr)
                    ? PptAnimationHelper.ParseEffectType(effectTypeStr)
                    : effect.Type;
                var newSubtype = !string.IsNullOrEmpty(effectSubtypeStr)
                    ? PptAnimationHelper.ParseEffectSubtype(effectSubtypeStr)
                    : effect.Subtype;
                var newTrigger = !string.IsNullOrEmpty(triggerTypeStr)
                    ? PptAnimationHelper.ParseTriggerType(triggerTypeStr)
                    : effect.Timing.TriggerType;

                var currentDuration = effect.Timing.Duration;
                var currentDelay = effect.Timing.TriggerDelayTime;

                sequence.Remove(effect);
                var newEffect = sequence.AddEffect(shape, newEffectType, newSubtype, newTrigger);
                newEffect.Timing.Duration = duration ?? currentDuration;
                newEffect.Timing.TriggerDelayTime = delay ?? currentDelay;
            }
        }
        else
        {
            foreach (var anim in animations)
                sequence.Remove(anim);

            if (!string.IsNullOrEmpty(effectTypeStr))
            {
                var effectType = PptAnimationHelper.ParseEffectType(effectTypeStr);
                var effectSubtype = PptAnimationHelper.ParseEffectSubtype(effectSubtypeStr);
                var triggerType = PptAnimationHelper.ParseTriggerType(triggerTypeStr);
                var effect = sequence.AddEffect(shape, effectType, effectSubtype, triggerType);

                if (duration.HasValue) effect.Timing.Duration = duration.Value;
                if (delay.HasValue) effect.Timing.TriggerDelayTime = delay.Value;
            }
        }

        MarkModified(context);

        return Success($"Animation updated on slide {slideIndex}, shape {shapeIndex}.");
    }
}
