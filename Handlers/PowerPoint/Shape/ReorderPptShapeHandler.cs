using Aspose.Slides;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Helpers.PowerPoint;
using AsposeMcpServer.Results.Common;

namespace AsposeMcpServer.Handlers.PowerPoint.Shape;

/// <summary>
///     Handler for reordering a shape's Z-order position.
/// </summary>
[ResultType(typeof(SuccessResult))]
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
    public override object Execute(OperationContext<Presentation> context, OperationParameters parameters)
    {
        var p = ExtractReorderPptShapeParameters(parameters);

        var presentation = context.Document;
        var slide = PowerPointHelper.GetSlide(presentation, p.SlideIndex);

        PowerPointHelper.ValidateCollectionIndex(p.ShapeIndex, slide.Shapes.Count, "shapeIndex");
        PowerPointHelper.ValidateCollectionIndex(p.ToIndex, slide.Shapes.Count, "toIndex");

        var shape = slide.Shapes[p.ShapeIndex];
        slide.Shapes.Reorder(p.ToIndex, shape);

        MarkModified(context);

        return new SuccessResult { Message = $"Shape Z-order changed: {p.ShapeIndex} -> {p.ToIndex}." };
    }

    /// <summary>
    ///     Extracts parameters for reorder shape operation.
    /// </summary>
    /// <param name="parameters">The operation parameters.</param>
    /// <returns>The extracted parameters.</returns>
    private static ReorderPptShapeParameters ExtractReorderPptShapeParameters(OperationParameters parameters)
    {
        return new ReorderPptShapeParameters(
            parameters.GetRequired<int>("slideIndex"),
            parameters.GetRequired<int>("shapeIndex"),
            parameters.GetRequired<int>("toIndex"));
    }

    /// <summary>
    ///     Parameters for reorder shape operation.
    /// </summary>
    /// <param name="SlideIndex">The slide index (0-based).</param>
    /// <param name="ShapeIndex">The shape index to reorder.</param>
    /// <param name="ToIndex">The target Z-order index.</param>
    private sealed record ReorderPptShapeParameters(int SlideIndex, int ShapeIndex, int ToIndex);
}
