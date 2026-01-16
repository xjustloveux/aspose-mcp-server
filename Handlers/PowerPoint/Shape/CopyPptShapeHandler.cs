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
        var p = ExtractCopyPptShapeParameters(parameters);
        var presentation = context.Document;

        PowerPointHelper.ValidateCollectionIndex(p.FromSlide, presentation.Slides.Count, "fromSlide");
        PowerPointHelper.ValidateCollectionIndex(p.ToSlide, presentation.Slides.Count, "toSlide");

        var sourceSlide = presentation.Slides[p.FromSlide];
        PowerPointHelper.ValidateCollectionIndex(p.ShapeIndex, sourceSlide.Shapes.Count, "shapeIndex");

        var targetSlide = presentation.Slides[p.ToSlide];
        targetSlide.Shapes.AddClone(sourceSlide.Shapes[p.ShapeIndex]);

        MarkModified(context);

        return Success($"Shape {p.ShapeIndex} copied from slide {p.FromSlide} to slide {p.ToSlide}.");
    }

    private static CopyPptShapeParameters ExtractCopyPptShapeParameters(OperationParameters parameters)
    {
        return new CopyPptShapeParameters(
            parameters.GetRequired<int>("fromSlide"),
            parameters.GetRequired<int>("toSlide"),
            parameters.GetRequired<int>("shapeIndex"));
    }

    private sealed record CopyPptShapeParameters(int FromSlide, int ToSlide, int ShapeIndex);
}
