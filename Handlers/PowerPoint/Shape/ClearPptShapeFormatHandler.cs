using Aspose.Slides;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Helpers;

namespace AsposeMcpServer.Handlers.PowerPoint.Shape;

/// <summary>
///     Handler for clearing shape format properties in PowerPoint presentations.
/// </summary>
public class ClearPptShapeFormatHandler : OperationHandlerBase<Presentation>
{
    /// <inheritdoc />
    public override string Operation => "clear_format";

    /// <summary>
    ///     Clears shape format properties (fill, line).
    /// </summary>
    /// <param name="context">The presentation context.</param>
    /// <param name="parameters">
    ///     Required: shapeIndex.
    ///     Optional: slideIndex (default: 0), clearFill (default: true), clearLine (default: true).
    /// </param>
    /// <returns>Success message with clear details.</returns>
    public override string Execute(OperationContext<Presentation> context, OperationParameters parameters)
    {
        var slideIndex = parameters.GetOptional("slideIndex", 0);
        var shapeIndex = parameters.GetRequired<int>("shapeIndex");
        var clearFill = parameters.GetOptional("clearFill", true);
        var clearLine = parameters.GetOptional("clearLine", true);

        var presentation = context.Document;
        var slide = PowerPointHelper.GetSlide(presentation, slideIndex);

        PowerPointHelper.ValidateCollectionIndex(shapeIndex, slide.Shapes.Count, "shapeIndex");

        var shape = slide.Shapes[shapeIndex];

        if (clearFill)
            shape.FillFormat.FillType = FillType.NoFill;

        if (clearLine)
            shape.LineFormat.FillFormat.FillType = FillType.NoFill;

        MarkModified(context);

        return Success($"Format cleared from shape {shapeIndex} on slide {slideIndex}.");
    }
}
