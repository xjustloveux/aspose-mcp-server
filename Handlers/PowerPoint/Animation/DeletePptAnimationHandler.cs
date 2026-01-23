using Aspose.Slides;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Helpers.PowerPoint;
using AsposeMcpServer.Results.Common;

namespace AsposeMcpServer.Handlers.PowerPoint.Animation;

/// <summary>
///     Handler for deleting animations from PowerPoint presentations.
/// </summary>
[ResultType(typeof(SuccessResult))]
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
    public override object Execute(OperationContext<Presentation> context, OperationParameters parameters)
    {
        var p = ExtractDeleteAnimationParameters(parameters);

        var presentation = context.Document;
        var slide = PowerPointHelper.GetSlide(presentation, p.SlideIndex);
        var sequence = slide.Timeline.MainSequence;

        if (p.ShapeIndex.HasValue)
        {
            PowerPointHelper.ValidateShapeIndex(p.ShapeIndex.Value, slide);
            var shape = slide.Shapes[p.ShapeIndex.Value];
            var animations = sequence.Where(e => e.TargetShape == shape).ToList();

            if (p.AnimationIndex.HasValue)
            {
                if (p.AnimationIndex.Value < 0 || p.AnimationIndex.Value >= animations.Count)
                    throw new ArgumentException($"animationIndex must be between 0 and {animations.Count - 1}");
                sequence.Remove(animations[p.AnimationIndex.Value]);
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

        return new SuccessResult { Message = $"Animation(s) deleted from slide {p.SlideIndex}." };
    }

    /// <summary>
    ///     Extracts delete animation parameters from operation parameters.
    /// </summary>
    /// <param name="parameters">The operation parameters.</param>
    /// <returns>The extracted delete animation parameters.</returns>
    private static DeleteAnimationParameters ExtractDeleteAnimationParameters(OperationParameters parameters)
    {
        return new DeleteAnimationParameters(
            parameters.GetRequired<int>("slideIndex"),
            parameters.GetOptional<int?>("shapeIndex"),
            parameters.GetOptional<int?>("animationIndex"));
    }

    /// <summary>
    ///     Record for holding delete animation parameters.
    /// </summary>
    /// <param name="SlideIndex">The slide index.</param>
    /// <param name="ShapeIndex">The optional shape index.</param>
    /// <param name="AnimationIndex">The optional animation index.</param>
    private sealed record DeleteAnimationParameters(int SlideIndex, int? ShapeIndex, int? AnimationIndex);
}
