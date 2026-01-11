using Aspose.Slides;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Helpers;

namespace AsposeMcpServer.Handlers.PowerPoint.Shape;

/// <summary>
///     Handler for grouping shapes in PowerPoint presentations.
/// </summary>
public class GroupPptShapesHandler : OperationHandlerBase<Presentation>
{
    /// <inheritdoc />
    public override string Operation => "group";

    /// <summary>
    ///     Groups multiple shapes into a single group shape.
    /// </summary>
    /// <param name="context">The presentation context.</param>
    /// <param name="parameters">
    ///     Required: shapeIndices (at least 2 shapes).
    ///     Optional: slideIndex (default: 0).
    /// </param>
    /// <returns>Success message with group details.</returns>
    public override string Execute(OperationContext<Presentation> context, OperationParameters parameters)
    {
        var slideIndex = parameters.GetOptional("slideIndex", 0);
        var shapeIndices = parameters.GetRequired<int[]>("shapeIndices");

        if (shapeIndices.Length < 2)
            throw new ArgumentException("At least 2 shapes are required for grouping");

        var presentation = context.Document;
        var slide = PowerPointHelper.GetSlide(presentation, slideIndex);

        foreach (var idx in shapeIndices)
            PowerPointHelper.ValidateCollectionIndex(idx, slide.Shapes.Count, "shapeIndex");

        var sortedIndices = shapeIndices.OrderBy(i => i).ToArray();
        var shapes = sortedIndices.Select(idx => slide.Shapes[idx]).ToList();

        var minX = shapes.Min(s => s.X);
        var minY = shapes.Min(s => s.Y);
        var maxX = shapes.Max(s => s.X + s.Width);
        var maxY = shapes.Max(s => s.Y + s.Height);

        var groupShape = slide.Shapes.AddGroupShape();
        groupShape.X = minX;
        groupShape.Y = minY;
        groupShape.Width = maxX - minX;
        groupShape.Height = maxY - minY;

        foreach (var idx in sortedIndices.Reverse())
        {
            var shape = slide.Shapes[idx];
            var clonedShape = groupShape.Shapes.AddClone(shape);
            clonedShape.X = shape.X - minX;
            clonedShape.Y = shape.Y - minY;
        }

        foreach (var idx in sortedIndices.Reverse())
            slide.Shapes.RemoveAt(idx);

        MarkModified(context);

        var newIndex = slide.Shapes.IndexOf(groupShape);
        return Success($"Grouped {shapeIndices.Length} shapes into group (shapeIndex: {newIndex}).");
    }
}
