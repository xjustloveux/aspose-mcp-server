using Aspose.Slides;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Helpers;

namespace AsposeMcpServer.Handlers.PowerPoint.Animation;

/// <summary>
///     Handler for adding animations to PowerPoint presentations.
/// </summary>
public class AddPptAnimationHandler : OperationHandlerBase<Presentation>
{
    /// <inheritdoc />
    public override string Operation => "add";

    /// <summary>
    ///     Adds animation to a shape.
    /// </summary>
    /// <param name="context">The presentation context.</param>
    /// <param name="parameters">
    ///     Required: slideIndex, shapeIndex
    ///     Optional: effectType, effectSubtype, triggerType
    /// </param>
    /// <returns>Success message with animation details.</returns>
    public override string Execute(OperationContext<Presentation> context, OperationParameters parameters)
    {
        var p = ExtractAddAnimationParameters(parameters);

        var presentation = context.Document;
        var slide = PowerPointHelper.GetSlide(presentation, p.SlideIndex);
        PowerPointHelper.ValidateShapeIndex(p.ShapeIndex, slide);
        var shape = slide.Shapes[p.ShapeIndex];

        var effectType = PptAnimationHelper.ParseEffectType(p.EffectType ?? "Fade");
        var effectSubtype = PptAnimationHelper.ParseEffectSubtype(p.EffectSubtype);
        var triggerType = PptAnimationHelper.ParseTriggerType(p.TriggerType);

        slide.Timeline.MainSequence.AddEffect(shape, effectType, effectSubtype, triggerType);

        MarkModified(context);

        return Success($"Animation '{p.EffectType ?? "Fade"}' added to shape {p.ShapeIndex} on slide {p.SlideIndex}.");
    }

    /// <summary>
    ///     Extracts add animation parameters from operation parameters.
    /// </summary>
    /// <param name="parameters">The operation parameters.</param>
    /// <returns>The extracted add animation parameters.</returns>
    private static AddAnimationParameters ExtractAddAnimationParameters(OperationParameters parameters)
    {
        return new AddAnimationParameters(
            parameters.GetRequired<int>("slideIndex"),
            parameters.GetRequired<int>("shapeIndex"),
            parameters.GetOptional<string?>("effectType"),
            parameters.GetOptional<string?>("effectSubtype"),
            parameters.GetOptional<string?>("triggerType"));
    }

    /// <summary>
    ///     Record for holding add animation parameters.
    /// </summary>
    /// <param name="SlideIndex">The slide index.</param>
    /// <param name="ShapeIndex">The shape index.</param>
    /// <param name="EffectType">The optional effect type.</param>
    /// <param name="EffectSubtype">The optional effect subtype.</param>
    /// <param name="TriggerType">The optional trigger type.</param>
    private record AddAnimationParameters(
        int SlideIndex,
        int ShapeIndex,
        string? EffectType,
        string? EffectSubtype,
        string? TriggerType);
}
