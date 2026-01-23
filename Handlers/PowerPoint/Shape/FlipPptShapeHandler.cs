using Aspose.Slides;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Helpers.PowerPoint;
using AsposeMcpServer.Results.Common;

namespace AsposeMcpServer.Handlers.PowerPoint.Shape;

/// <summary>
///     Handler for flipping a shape horizontally or vertically.
/// </summary>
[ResultType(typeof(SuccessResult))]
public class FlipPptShapeHandler : OperationHandlerBase<Presentation>
{
    /// <inheritdoc />
    public override string Operation => "flip";

    /// <summary>
    ///     Flips a shape horizontally or vertically.
    /// </summary>
    /// <param name="context">The presentation context.</param>
    /// <param name="parameters">
    ///     Required: slideIndex, shapeIndex
    ///     Required: at least one of flipHorizontal or flipVertical must be provided
    /// </param>
    /// <returns>Success message with flip details.</returns>
    public override object Execute(OperationContext<Presentation> context, OperationParameters parameters)
    {
        var p = ExtractFlipPptShapeParameters(parameters);

        if (!p.FlipHorizontal.HasValue && !p.FlipVertical.HasValue)
            throw new ArgumentException("At least one of flipHorizontal or flipVertical must be provided");

        var presentation = context.Document;
        var slide = PowerPointHelper.GetSlide(presentation, p.SlideIndex);
        var shape = PowerPointHelper.GetShape(slide, p.ShapeIndex);

        var currentFrame = shape.Frame;
        var newFlipH = GetFlipValue(p.FlipHorizontal, currentFrame.FlipH);
        var newFlipV = GetFlipValue(p.FlipVertical, currentFrame.FlipV);

        shape.Frame = new ShapeFrame(
            currentFrame.X, currentFrame.Y, currentFrame.Width, currentFrame.Height,
            newFlipH, newFlipV, currentFrame.Rotation);

        MarkModified(context);

        List<string> flipDesc = [];
        if (p.FlipHorizontal.HasValue) flipDesc.Add($"H={p.FlipHorizontal.Value}");
        if (p.FlipVertical.HasValue) flipDesc.Add($"V={p.FlipVertical.Value}");

        return new SuccessResult { Message = $"Shape {p.ShapeIndex} flipped ({string.Join(", ", flipDesc)})." };
    }

    /// <summary>
    ///     Gets the flip value from nullable bool or uses current value.
    /// </summary>
    /// <param name="flipValue">The nullable flip value from parameters.</param>
    /// <param name="currentValue">The current flip value from the frame.</param>
    /// <returns>The NullableBool flip value to use.</returns>
    private static NullableBool GetFlipValue(bool? flipValue, NullableBool currentValue)
    {
        if (!flipValue.HasValue) return currentValue;
        return flipValue.Value ? NullableBool.True : NullableBool.False;
    }

    /// <summary>
    ///     Extracts parameters for flip shape operation.
    /// </summary>
    /// <param name="parameters">The operation parameters.</param>
    /// <returns>The extracted parameters.</returns>
    private static FlipPptShapeParameters ExtractFlipPptShapeParameters(OperationParameters parameters)
    {
        return new FlipPptShapeParameters(
            parameters.GetRequired<int>("slideIndex"),
            parameters.GetRequired<int>("shapeIndex"),
            parameters.GetOptional<bool?>("flipHorizontal"),
            parameters.GetOptional<bool?>("flipVertical"));
    }

    /// <summary>
    ///     Parameters for flip shape operation.
    /// </summary>
    /// <param name="SlideIndex">The slide index (0-based).</param>
    /// <param name="ShapeIndex">The shape index to flip.</param>
    /// <param name="FlipHorizontal">Whether to flip horizontally.</param>
    /// <param name="FlipVertical">Whether to flip vertically.</param>
    private sealed record FlipPptShapeParameters(
        int SlideIndex,
        int ShapeIndex,
        bool? FlipHorizontal,
        bool? FlipVertical);
}
