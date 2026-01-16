using Aspose.Slides;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Helpers;

namespace AsposeMcpServer.Handlers.PowerPoint.Shape;

/// <summary>
///     Handler for editing shape properties in PowerPoint presentations.
/// </summary>
public class EditPptShapeHandler : OperationHandlerBase<Presentation>
{
    /// <inheritdoc />
    public override string Operation => "edit";

    /// <summary>
    ///     Edits shape properties such as position, size, rotation, and text.
    /// </summary>
    /// <param name="context">The presentation context.</param>
    /// <param name="parameters">
    ///     Required: shapeIndex.
    ///     Optional: slideIndex (default: 0), x, y, width, height, rotation, name, text.
    /// </param>
    /// <returns>Success message with edit details.</returns>
    public override string Execute(OperationContext<Presentation> context, OperationParameters parameters)
    {
        var p = ExtractEditPptShapeParameters(parameters);

        var presentation = context.Document;
        var slide = PowerPointHelper.GetSlide(presentation, p.SlideIndex);

        PowerPointHelper.ValidateCollectionIndex(p.ShapeIndex, slide.Shapes.Count, "shapeIndex");

        var shape = slide.Shapes[p.ShapeIndex];

        if (p.X.HasValue)
            shape.X = p.X.Value;

        if (p.Y.HasValue)
            shape.Y = p.Y.Value;

        if (p.Width.HasValue)
            shape.Width = p.Width.Value;

        if (p.Height.HasValue)
            shape.Height = p.Height.Value;

        if (p.Rotation.HasValue)
            shape.Rotation = p.Rotation.Value;

        if (!string.IsNullOrEmpty(p.Name))
            shape.Name = p.Name;

        if (!string.IsNullOrEmpty(p.Text) && shape is IAutoShape { TextFrame: not null } autoShape)
            autoShape.TextFrame.Text = p.Text;

        MarkModified(context);

        return Success($"Shape {p.ShapeIndex} updated on slide {p.SlideIndex}.");
    }

    /// <summary>
    ///     Extracts parameters for edit shape operation.
    /// </summary>
    /// <param name="parameters">The operation parameters.</param>
    /// <returns>The extracted parameters.</returns>
    private static EditPptShapeParameters ExtractEditPptShapeParameters(OperationParameters parameters)
    {
        return new EditPptShapeParameters(
            parameters.GetOptional("slideIndex", 0),
            parameters.GetRequired<int>("shapeIndex"),
            parameters.GetOptional<float?>("x"),
            parameters.GetOptional<float?>("y"),
            parameters.GetOptional<float?>("width"),
            parameters.GetOptional<float?>("height"),
            parameters.GetOptional<float?>("rotation"),
            parameters.GetOptional<string?>("name"),
            parameters.GetOptional<string?>("text"));
    }

    /// <summary>
    ///     Parameters for edit shape operation.
    /// </summary>
    /// <param name="SlideIndex">The slide index (0-based).</param>
    /// <param name="ShapeIndex">The shape index to edit.</param>
    /// <param name="X">The X position.</param>
    /// <param name="Y">The Y position.</param>
    /// <param name="Width">The width.</param>
    /// <param name="Height">The height.</param>
    /// <param name="Rotation">The rotation angle.</param>
    /// <param name="Name">The shape name.</param>
    /// <param name="Text">The shape text.</param>
    private sealed record EditPptShapeParameters(
        int SlideIndex,
        int ShapeIndex,
        float? X,
        float? Y,
        float? Width,
        float? Height,
        float? Rotation,
        string? Name,
        string? Text);
}
