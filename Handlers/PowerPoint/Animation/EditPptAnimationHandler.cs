using Aspose.Slides;
using Aspose.Slides.Animation;
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
        var animParams = ExtractAnimationParameters(parameters);
        var presentation = context.Document;
        var slide = PowerPointHelper.GetSlide(presentation, animParams.SlideIndex);
        PowerPointHelper.ValidateShapeIndex(animParams.ShapeIndex, slide);

        var shape = slide.Shapes[animParams.ShapeIndex];
        var sequence = slide.Timeline.MainSequence;
        var animations = sequence.Where(e => e.TargetShape == shape).ToList();

        if (animParams.AnimationIndex.HasValue)
            EditSpecificAnimation(sequence, shape, animations, animParams);
        else
            ReplaceAllAnimations(sequence, shape, animations, animParams);

        MarkModified(context);
        return Success($"Animation updated on slide {animParams.SlideIndex}, shape {animParams.ShapeIndex}.");
    }

    /// <summary>
    ///     Extracts animation parameters from operation parameters.
    /// </summary>
    /// <param name="parameters">The operation parameters.</param>
    /// <returns>The extracted animation parameters.</returns>
    private static AnimationParameters ExtractAnimationParameters(OperationParameters parameters)
    {
        return new AnimationParameters(
            parameters.GetRequired<int>("slideIndex"),
            parameters.GetRequired<int>("shapeIndex"),
            parameters.GetOptional<int?>("animationIndex"),
            parameters.GetOptional<string?>("effectType"),
            parameters.GetOptional<string?>("effectSubtype"),
            parameters.GetOptional<string?>("triggerType"),
            parameters.GetOptional<float?>("duration"),
            parameters.GetOptional<float?>("delay")
        );
    }

    /// <summary>
    ///     Edits a specific animation by index.
    /// </summary>
    /// <param name="sequence">The animation sequence.</param>
    /// <param name="shape">The target shape.</param>
    /// <param name="animations">The list of animations for the shape.</param>
    /// <param name="p">The animation parameters.</param>
    private static void EditSpecificAnimation(ISequence sequence, IShape shape,
        List<IEffect> animations, AnimationParameters p)
    {
        if (p.AnimationIndex!.Value < 0 || p.AnimationIndex.Value >= animations.Count)
            throw new ArgumentException($"animationIndex must be between 0 and {animations.Count - 1}");

        var effect = animations[p.AnimationIndex.Value];
        ApplyTimingIfProvided(effect, p.Duration, p.Delay);

        if (ShouldRecreateEffect(p))
            RecreateEffect(sequence, shape, effect, p);
    }

    /// <summary>
    ///     Determines whether the effect needs to be recreated based on parameters.
    /// </summary>
    /// <param name="p">The animation parameters.</param>
    /// <returns>True if the effect should be recreated, false otherwise.</returns>
    private static bool ShouldRecreateEffect(AnimationParameters p)
    {
        return !string.IsNullOrEmpty(p.EffectType) || !string.IsNullOrEmpty(p.EffectSubtype) ||
               !string.IsNullOrEmpty(p.TriggerType);
    }

    /// <summary>
    ///     Recreates an animation effect with new parameters.
    /// </summary>
    /// <param name="sequence">The animation sequence.</param>
    /// <param name="shape">The target shape.</param>
    /// <param name="effect">The original effect to recreate.</param>
    /// <param name="p">The animation parameters.</param>
    private static void RecreateEffect(ISequence sequence, IShape shape, IEffect effect, AnimationParameters p)
    {
        var newEffectType = !string.IsNullOrEmpty(p.EffectType)
            ? PptAnimationHelper.ParseEffectType(p.EffectType)
            : effect.Type;
        var newSubtype = !string.IsNullOrEmpty(p.EffectSubtype)
            ? PptAnimationHelper.ParseEffectSubtype(p.EffectSubtype)
            : effect.Subtype;
        var newTrigger = !string.IsNullOrEmpty(p.TriggerType)
            ? PptAnimationHelper.ParseTriggerType(p.TriggerType)
            : effect.Timing.TriggerType;

        var currentDuration = effect.Timing.Duration;
        var currentDelay = effect.Timing.TriggerDelayTime;

        sequence.Remove(effect);
        var newEffect = sequence.AddEffect(shape, newEffectType, newSubtype, newTrigger);
        newEffect.Timing.Duration = p.Duration ?? currentDuration;
        newEffect.Timing.TriggerDelayTime = p.Delay ?? currentDelay;
    }

    /// <summary>
    ///     Replaces all animations for a shape with a new animation.
    /// </summary>
    /// <param name="sequence">The animation sequence.</param>
    /// <param name="shape">The target shape.</param>
    /// <param name="animations">The list of existing animations to remove.</param>
    /// <param name="p">The animation parameters for the new animation.</param>
    private static void ReplaceAllAnimations(ISequence sequence, IShape shape,
        List<IEffect> animations, AnimationParameters p)
    {
        foreach (var anim in animations)
            sequence.Remove(anim);

        if (!string.IsNullOrEmpty(p.EffectType))
        {
            var effect = sequence.AddEffect(shape,
                PptAnimationHelper.ParseEffectType(p.EffectType),
                PptAnimationHelper.ParseEffectSubtype(p.EffectSubtype),
                PptAnimationHelper.ParseTriggerType(p.TriggerType));
            ApplyTimingIfProvided(effect, p.Duration, p.Delay);
        }
    }

    /// <summary>
    ///     Applies timing settings to an effect if provided.
    /// </summary>
    /// <param name="effect">The effect to apply timing to.</param>
    /// <param name="duration">The optional duration value.</param>
    /// <param name="delay">The optional delay value.</param>
    private static void ApplyTimingIfProvided(IEffect effect, float? duration, float? delay)
    {
        if (duration.HasValue) effect.Timing.Duration = duration.Value;
        if (delay.HasValue) effect.Timing.TriggerDelayTime = delay.Value;
    }

    /// <summary>
    ///     Record for holding animation parameters extracted from operation parameters.
    /// </summary>
    /// <param name="SlideIndex">The slide index.</param>
    /// <param name="ShapeIndex">The shape index.</param>
    /// <param name="AnimationIndex">The optional animation index.</param>
    /// <param name="EffectType">The optional effect type.</param>
    /// <param name="EffectSubtype">The optional effect subtype.</param>
    /// <param name="TriggerType">The optional trigger type.</param>
    /// <param name="Duration">The optional duration.</param>
    /// <param name="Delay">The optional delay.</param>
    private record AnimationParameters(
        int SlideIndex,
        int ShapeIndex,
        int? AnimationIndex,
        string? EffectType,
        string? EffectSubtype,
        string? TriggerType,
        float? Duration,
        float? Delay);
}
