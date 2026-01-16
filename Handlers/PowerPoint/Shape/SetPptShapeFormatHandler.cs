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
        var p = ExtractSetPptShapeFormatParameters(parameters);

        var presentation = context.Document;
        var slide = PowerPointHelper.GetSlide(presentation, p.SlideIndex);

        PowerPointHelper.ValidateCollectionIndex(p.ShapeIndex, slide.Shapes.Count, "shapeIndex");

        var shape = slide.Shapes[p.ShapeIndex];

        if (!string.IsNullOrEmpty(p.FillColor))
        {
            shape.FillFormat.FillType = FillType.Solid;
            shape.FillFormat.SolidFillColor.Color = ColorHelper.ParseColor(p.FillColor);
        }

        if (!string.IsNullOrEmpty(p.LineColor))
        {
            shape.LineFormat.FillFormat.FillType = FillType.Solid;
            shape.LineFormat.FillFormat.SolidFillColor.Color = ColorHelper.ParseColor(p.LineColor);
        }

        if (p.LineWidth.HasValue)
            shape.LineFormat.Width = p.LineWidth.Value;

        if (p.Transparency.HasValue && shape.FillFormat.FillType == FillType.Solid)
        {
            var color = shape.FillFormat.SolidFillColor.Color;
            var alpha = (int)((1 - p.Transparency.Value) * 255);
            shape.FillFormat.SolidFillColor.Color = Color.FromArgb(alpha, color);
        }

        MarkModified(context);

        return Success($"Format applied to shape {p.ShapeIndex} on slide {p.SlideIndex}.");
    }

    /// <summary>
    ///     Extracts parameters for set shape format operation.
    /// </summary>
    /// <param name="parameters">The operation parameters.</param>
    /// <returns>The extracted parameters.</returns>
    private static SetPptShapeFormatParameters ExtractSetPptShapeFormatParameters(OperationParameters parameters)
    {
        return new SetPptShapeFormatParameters(
            parameters.GetOptional("slideIndex", 0),
            parameters.GetRequired<int>("shapeIndex"),
            parameters.GetOptional<string?>("fillColor"),
            parameters.GetOptional<string?>("lineColor"),
            parameters.GetOptional<float?>("lineWidth"),
            parameters.GetOptional<float?>("transparency"));
    }

    /// <summary>
    ///     Parameters for set shape format operation.
    /// </summary>
    /// <param name="SlideIndex">The slide index (0-based).</param>
    /// <param name="ShapeIndex">The shape index.</param>
    /// <param name="FillColor">The fill color.</param>
    /// <param name="LineColor">The line color.</param>
    /// <param name="LineWidth">The line width.</param>
    /// <param name="Transparency">The transparency value (0-1).</param>
    private sealed record SetPptShapeFormatParameters(
        int SlideIndex,
        int ShapeIndex,
        string? FillColor,
        string? LineColor,
        float? LineWidth,
        float? Transparency);
}
