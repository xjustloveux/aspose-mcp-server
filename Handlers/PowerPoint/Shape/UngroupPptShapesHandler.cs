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
        var p = ExtractUngroupPptShapesParameters(parameters);

        var presentation = context.Document;
        var slide = PowerPointHelper.GetSlide(presentation, p.SlideIndex);
        var shape = PowerPointHelper.GetShape(slide, p.ShapeIndex);

        if (shape is not IGroupShape groupShape)
            throw new ArgumentException($"Shape at index {p.ShapeIndex} is not a group");

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

        return Success($"Ungrouped {shapesInGroup.Count} shapes on slide {p.SlideIndex}.");
    }

    /// <summary>
    ///     Extracts parameters for ungroup shapes operation.
    /// </summary>
    /// <param name="parameters">The operation parameters.</param>
    /// <returns>The extracted parameters.</returns>
    private static UngroupPptShapesParameters ExtractUngroupPptShapesParameters(OperationParameters parameters)
    {
        return new UngroupPptShapesParameters(
            parameters.GetRequired<int>("slideIndex"),
            parameters.GetRequired<int>("shapeIndex"));
    }

    /// <summary>
    ///     Parameters for ungroup shapes operation.
    /// </summary>
    /// <param name="SlideIndex">The slide index (0-based).</param>
    /// <param name="ShapeIndex">The group shape index to ungroup.</param>
    private sealed record UngroupPptShapesParameters(int SlideIndex, int ShapeIndex);
}
