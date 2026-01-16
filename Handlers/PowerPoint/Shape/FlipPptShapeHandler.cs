using Aspose.Slides;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Helpers;

namespace AsposeMcpServer.Handlers.PowerPoint.Shape;

/// <summary>
///     Handler for flipping a shape horizontally or vertically.
/// </summary>
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
    public override string Execute(OperationContext<Presentation> context, OperationParameters parameters)
    {
        var p = ExtractFlipPptShapeParameters(parameters);

        if (!p.FlipHorizontal.HasValue && !p.FlipVertical.HasValue)
            throw new ArgumentException("At least one of flipHorizontal or flipVertical must be provided");

        var presentation = context.Document;
        var slide = PowerPointHelper.GetSlide(presentation, p.SlideIndex);
        var shape = PowerPointHelper.GetShape(slide, p.ShapeIndex);

        var currentFrame = shape.Frame;
        var newFlipH = p.FlipHorizontal.HasValue
            ? p.FlipHorizontal.Value ? NullableBool.True : NullableBool.False
            : currentFrame.FlipH;
        var newFlipV = p.FlipVertical.HasValue
            ? p.FlipVertical.Value ? NullableBool.True : NullableBool.False
            : currentFrame.FlipV;

        shape.Frame = new ShapeFrame(
            currentFrame.X, currentFrame.Y, currentFrame.Width, currentFrame.Height,
            newFlipH, newFlipV, currentFrame.Rotation);

        MarkModified(context);

        List<string> flipDesc = [];
        if (p.FlipHorizontal.HasValue) flipDesc.Add($"H={p.FlipHorizontal.Value}");
        if (p.FlipVertical.HasValue) flipDesc.Add($"V={p.FlipVertical.Value}");

        return Success($"Shape {p.ShapeIndex} flipped ({string.Join(", ", flipDesc)}).");
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
    private record FlipPptShapeParameters(int SlideIndex, int ShapeIndex, bool? FlipHorizontal, bool? FlipVertical);
}
