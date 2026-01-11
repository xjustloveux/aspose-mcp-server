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
        var slideIndex = parameters.GetRequired<int>("slideIndex");
        var shapeIndex = parameters.GetRequired<int>("shapeIndex");
        var effectTypeStr = parameters.GetOptional<string?>("effectType");
        var effectSubtypeStr = parameters.GetOptional<string?>("effectSubtype");
        var triggerTypeStr = parameters.GetOptional<string?>("triggerType");

        var presentation = context.Document;
        var slide = PowerPointHelper.GetSlide(presentation, slideIndex);
        PowerPointHelper.ValidateShapeIndex(shapeIndex, slide);
        var shape = slide.Shapes[shapeIndex];

        var effectType = PptAnimationHelper.ParseEffectType(effectTypeStr ?? "Fade");
        var effectSubtype = PptAnimationHelper.ParseEffectSubtype(effectSubtypeStr);
        var triggerType = PptAnimationHelper.ParseTriggerType(triggerTypeStr);

        slide.Timeline.MainSequence.AddEffect(shape, effectType, effectSubtype, triggerType);

        MarkModified(context);

        return Success($"Animation '{effectTypeStr ?? "Fade"}' added to shape {shapeIndex} on slide {slideIndex}.");
    }
}
