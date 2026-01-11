using Aspose.Slides;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Helpers;

namespace AsposeMcpServer.Handlers.PowerPoint.Shape;

/// <summary>
///     Handler for copying a shape to another slide.
/// </summary>
public class CopyPptShapeHandler : OperationHandlerBase<Presentation>
{
    /// <inheritdoc />
    public override string Operation => "copy";

    /// <summary>
    ///     Copies a shape to another slide.
    /// </summary>
    /// <param name="context">The presentation context.</param>
    /// <param name="parameters">
    ///     Required: fromSlide, toSlide, shapeIndex
    /// </param>
    /// <returns>Success message with copy details.</returns>
    public override string Execute(OperationContext<Presentation> context, OperationParameters parameters)
    {
        var fromSlide = parameters.GetRequired<int>("fromSlide");
        var toSlide = parameters.GetRequired<int>("toSlide");
        var shapeIndex = parameters.GetRequired<int>("shapeIndex");

        var presentation = context.Document;

        PowerPointHelper.ValidateCollectionIndex(fromSlide, presentation.Slides.Count, "fromSlide");
        PowerPointHelper.ValidateCollectionIndex(toSlide, presentation.Slides.Count, "toSlide");

        var sourceSlide = presentation.Slides[fromSlide];
        PowerPointHelper.ValidateCollectionIndex(shapeIndex, sourceSlide.Shapes.Count, "shapeIndex");

        var targetSlide = presentation.Slides[toSlide];
        targetSlide.Shapes.AddClone(sourceSlide.Shapes[shapeIndex]);

        MarkModified(context);

        return Success($"Shape {shapeIndex} copied from slide {fromSlide} to slide {toSlide}.");
    }
}
