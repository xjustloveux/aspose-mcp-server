using System.Drawing;
using Aspose.Slides;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Helpers;

namespace AsposeMcpServer.Handlers.PowerPoint.Shape;

/// <summary>
///     Handler for setting shape format properties in PowerPoint presentations.
/// </summary>
public class SetPptShapeFormatHandler : OperationHandlerBase<Presentation>
{
    /// <inheritdoc />
    public override string Operation => "set_format";

    /// <summary>
    ///     Sets shape format properties such as fill color, line color, and line width.
    /// </summary>
    /// <param name="context">The presentation context.</param>
    /// <param name="parameters">
    ///     Required: shapeIndex.
    ///     Optional: slideIndex (default: 0), fillColor, lineColor, lineWidth, transparency.
    /// </param>
    /// <returns>Success message with format details.</returns>
    public override string Execute(OperationContext<Presentation> context, OperationParameters parameters)
    {
        var slideIndex = parameters.GetOptional("slideIndex", 0);
        var shapeIndex = parameters.GetRequired<int>("shapeIndex");
        var fillColor = parameters.GetOptional<string?>("fillColor");
        var lineColor = parameters.GetOptional<string?>("lineColor");
        var lineWidth = parameters.GetOptional<float?>("lineWidth");
        var transparency = parameters.GetOptional<float?>("transparency");

        var presentation = context.Document;
        var slide = PowerPointHelper.GetSlide(presentation, slideIndex);

        PowerPointHelper.ValidateCollectionIndex(shapeIndex, slide.Shapes.Count, "shapeIndex");

        var shape = slide.Shapes[shapeIndex];

        if (!string.IsNullOrEmpty(fillColor))
        {
            shape.FillFormat.FillType = FillType.Solid;
            shape.FillFormat.SolidFillColor.Color = ColorHelper.ParseColor(fillColor);
        }

        if (!string.IsNullOrEmpty(lineColor))
        {
            shape.LineFormat.FillFormat.FillType = FillType.Solid;
            shape.LineFormat.FillFormat.SolidFillColor.Color = ColorHelper.ParseColor(lineColor);
        }

        if (lineWidth.HasValue)
            shape.LineFormat.Width = lineWidth.Value;

        if (transparency.HasValue && shape.FillFormat.FillType == FillType.Solid)
        {
            var color = shape.FillFormat.SolidFillColor.Color;
            var alpha = (int)((1 - transparency.Value) * 255);
            shape.FillFormat.SolidFillColor.Color = Color.FromArgb(alpha, color);
        }

        MarkModified(context);

        return Success($"Format applied to shape {shapeIndex} on slide {slideIndex}.");
    }
}
