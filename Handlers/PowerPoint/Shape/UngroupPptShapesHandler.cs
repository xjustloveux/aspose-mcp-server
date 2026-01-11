using Aspose.Slides;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Helpers;

namespace AsposeMcpServer.Handlers.PowerPoint.Shape;

/// <summary>
///     Handler for ungrouping a group shape.
/// </summary>
public class UngroupPptShapesHandler : OperationHandlerBase<Presentation>
{
    /// <inheritdoc />
    public override string Operation => "ungroup";

    /// <summary>
    ///     Ungroups a group shape.
    /// </summary>
    /// <param name="context">The presentation context.</param>
    /// <param name="parameters">
    ///     Required: slideIndex, shapeIndex
    /// </param>
    /// <returns>Success message with ungroup details.</returns>
    public override string Execute(OperationContext<Presentation> context, OperationParameters parameters)
    {
        var slideIndex = parameters.GetRequired<int>("slideIndex");
        var shapeIndex = parameters.GetRequired<int>("shapeIndex");

        var presentation = context.Document;
        var slide = PowerPointHelper.GetSlide(presentation, slideIndex);
        var shape = PowerPointHelper.GetShape(slide, shapeIndex);

        if (shape is not IGroupShape groupShape)
            throw new ArgumentException($"Shape at index {shapeIndex} is not a group");

        var groupIndex = slide.Shapes.IndexOf(groupShape);
        var shapesInGroup = groupShape.Shapes.ToList();

        var insertIndex = groupIndex;
        foreach (var s in shapesInGroup)
        {
            slide.Shapes.InsertClone(insertIndex, s);
            insertIndex++;
        }

        slide.Shapes.Remove(groupShape);

        MarkModified(context);

        return Success($"Ungrouped {shapesInGroup.Count} shapes on slide {slideIndex}.");
    }
}
