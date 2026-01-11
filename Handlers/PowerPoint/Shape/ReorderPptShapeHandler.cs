using Aspose.Slides;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Helpers;

namespace AsposeMcpServer.Handlers.PowerPoint.Shape;

/// <summary>
///     Handler for reordering a shape's Z-order position.
/// </summary>
public class ReorderPptShapeHandler : OperationHandlerBase<Presentation>
{
    /// <inheritdoc />
    public override string Operation => "reorder";

    /// <summary>
    ///     Reorders a shape's Z-order position.
    /// </summary>
    /// <param name="context">The presentation context.</param>
    /// <param name="parameters">
    ///     Required: slideIndex, shapeIndex, toIndex
    /// </param>
    /// <returns>Success message with reorder details.</returns>
    public override string Execute(OperationContext<Presentation> context, OperationParameters parameters)
    {
        var slideIndex = parameters.GetRequired<int>("slideIndex");
        var shapeIndex = parameters.GetRequired<int>("shapeIndex");
        var toIndex = parameters.GetRequired<int>("toIndex");

        var presentation = context.Document;
        var slide = PowerPointHelper.GetSlide(presentation, slideIndex);

        PowerPointHelper.ValidateCollectionIndex(shapeIndex, slide.Shapes.Count, "shapeIndex");
        PowerPointHelper.ValidateCollectionIndex(toIndex, slide.Shapes.Count, "toIndex");

        var shape = slide.Shapes[shapeIndex];
        slide.Shapes.Reorder(toIndex, shape);

        MarkModified(context);

        return Success($"Shape Z-order changed: {shapeIndex} -> {toIndex}.");
    }
}
