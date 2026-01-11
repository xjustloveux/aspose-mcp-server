using Aspose.Slides;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Helpers;

namespace AsposeMcpServer.Handlers.PowerPoint.Animation;

/// <summary>
///     Handler for deleting animations from PowerPoint presentations.
/// </summary>
public class DeletePptAnimationHandler : OperationHandlerBase<Presentation>
{
    /// <inheritdoc />
    public override string Operation => "delete";

    /// <summary>
    ///     Deletes animation(s) from a shape or slide.
    /// </summary>
    /// <param name="context">The presentation context.</param>
    /// <param name="parameters">
    ///     Required: slideIndex
    ///     Optional: shapeIndex, animationIndex
    /// </param>
    /// <returns>Success message with deletion details.</returns>
    public override string Execute(OperationContext<Presentation> context, OperationParameters parameters)
    {
        var slideIndex = parameters.GetRequired<int>("slideIndex");
        var shapeIndex = parameters.GetOptional<int?>("shapeIndex");
        var animationIndex = parameters.GetOptional<int?>("animationIndex");

        var presentation = context.Document;
        var slide = PowerPointHelper.GetSlide(presentation, slideIndex);
        var sequence = slide.Timeline.MainSequence;

        if (shapeIndex.HasValue)
        {
            PowerPointHelper.ValidateShapeIndex(shapeIndex.Value, slide);
            var shape = slide.Shapes[shapeIndex.Value];
            var animations = sequence.Where(e => e.TargetShape == shape).ToList();

            if (animationIndex.HasValue)
            {
                if (animationIndex.Value < 0 || animationIndex.Value >= animations.Count)
                    throw new ArgumentException($"animationIndex must be between 0 and {animations.Count - 1}");
                sequence.Remove(animations[animationIndex.Value]);
            }
            else
            {
                foreach (var anim in animations)
                    sequence.Remove(anim);
            }
        }
        else
        {
            sequence.Clear();
        }

        MarkModified(context);

        return Success($"Animation(s) deleted from slide {slideIndex}.");
    }
}
