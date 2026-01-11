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
        var slideIndex = parameters.GetRequired<int>("slideIndex");
        var shapeIndex = parameters.GetRequired<int>("shapeIndex");
        var flipHorizontal = parameters.GetOptional<bool?>("flipHorizontal");
        var flipVertical = parameters.GetOptional<bool?>("flipVertical");

        if (!flipHorizontal.HasValue && !flipVertical.HasValue)
            throw new ArgumentException("At least one of flipHorizontal or flipVertical must be provided");

        var presentation = context.Document;
        var slide = PowerPointHelper.GetSlide(presentation, slideIndex);
        var shape = PowerPointHelper.GetShape(slide, shapeIndex);

        var currentFrame = shape.Frame;
        var newFlipH = flipHorizontal.HasValue
            ? flipHorizontal.Value ? NullableBool.True : NullableBool.False
            : currentFrame.FlipH;
        var newFlipV = flipVertical.HasValue
            ? flipVertical.Value ? NullableBool.True : NullableBool.False
            : currentFrame.FlipV;

        shape.Frame = new ShapeFrame(
            currentFrame.X, currentFrame.Y, currentFrame.Width, currentFrame.Height,
            newFlipH, newFlipV, currentFrame.Rotation);

        MarkModified(context);

        List<string> flipDesc = [];
        if (flipHorizontal.HasValue) flipDesc.Add($"H={flipHorizontal.Value}");
        if (flipVertical.HasValue) flipDesc.Add($"V={flipVertical.Value}");

        return Success($"Shape {shapeIndex} flipped ({string.Join(", ", flipDesc)}).");
    }
}
