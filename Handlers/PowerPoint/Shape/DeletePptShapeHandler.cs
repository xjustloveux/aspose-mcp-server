using Aspose.Slides;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Helpers.PowerPoint;
using AsposeMcpServer.Results.Common;

namespace AsposeMcpServer.Handlers.PowerPoint.Shape;

/// <summary>
///     Handler for deleting shapes from PowerPoint slides.
/// </summary>
[ResultType(typeof(SuccessResult))]
public class DeletePptShapeHandler : OperationHandlerBase<Presentation>
{
    /// <inheritdoc />
    public override string Operation => "delete";

    /// <summary>
    ///     Deletes a shape from a slide.
    /// </summary>
    /// <param name="context">The presentation context.</param>
    /// <param name="parameters">
    ///     Required: shapeIndex.
    ///     Optional: slideIndex (default: 0).
    /// </param>
    /// <returns>Success message with deletion details.</returns>
    public override object Execute(OperationContext<Presentation> context, OperationParameters parameters)
    {
        var p = ExtractDeletePptShapeParameters(parameters);

        var presentation = context.Document;
        var slide = PowerPointHelper.GetSlide(presentation, p.SlideIndex);

        PowerPointHelper.ValidateCollectionIndex(p.ShapeIndex, slide.Shapes.Count, "shapeIndex");

        slide.Shapes.RemoveAt(p.ShapeIndex);

        MarkModified(context);

        return new SuccessResult { Message = $"Shape {p.ShapeIndex} deleted from slide {p.SlideIndex}." };
    }

    /// <summary>
    ///     Extracts parameters for delete shape operation.
    /// </summary>
    /// <param name="parameters">The operation parameters.</param>
    /// <returns>The extracted parameters.</returns>
    private static DeletePptShapeParameters ExtractDeletePptShapeParameters(OperationParameters parameters)
    {
        return new DeletePptShapeParameters(
            parameters.GetOptional("slideIndex", 0),
            parameters.GetRequired<int>("shapeIndex"));
    }

    /// <summary>
    ///     Parameters for delete shape operation.
    /// </summary>
    /// <param name="SlideIndex">The slide index (0-based).</param>
    /// <param name="ShapeIndex">The shape index to delete.</param>
    private sealed record DeletePptShapeParameters(int SlideIndex, int ShapeIndex);
}
