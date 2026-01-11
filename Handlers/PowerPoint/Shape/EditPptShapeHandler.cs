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
        var slideIndex = parameters.GetOptional("slideIndex", 0);
        var shapeIndex = parameters.GetRequired<int>("shapeIndex");
        var x = parameters.GetOptional<float?>("x");
        var y = parameters.GetOptional<float?>("y");
        var width = parameters.GetOptional<float?>("width");
        var height = parameters.GetOptional<float?>("height");
        var rotation = parameters.GetOptional<float?>("rotation");
        var name = parameters.GetOptional<string?>("name");
        var text = parameters.GetOptional<string?>("text");

        var presentation = context.Document;
        var slide = PowerPointHelper.GetSlide(presentation, slideIndex);

        PowerPointHelper.ValidateCollectionIndex(shapeIndex, slide.Shapes.Count, "shapeIndex");

        var shape = slide.Shapes[shapeIndex];

        if (x.HasValue)
            shape.X = x.Value;

        if (y.HasValue)
            shape.Y = y.Value;

        if (width.HasValue)
            shape.Width = width.Value;

        if (height.HasValue)
            shape.Height = height.Value;

        if (rotation.HasValue)
            shape.Rotation = rotation.Value;

        if (!string.IsNullOrEmpty(name))
            shape.Name = name;

        if (!string.IsNullOrEmpty(text) && shape is IAutoShape { TextFrame: not null } autoShape)
            autoShape.TextFrame.Text = text;

        MarkModified(context);

        return Success($"Shape {shapeIndex} updated on slide {slideIndex}.");
    }
}
