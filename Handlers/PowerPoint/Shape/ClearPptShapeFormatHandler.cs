using Aspose.Slides;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Helpers.PowerPoint;
using AsposeMcpServer.Results.Common;

namespace AsposeMcpServer.Handlers.PowerPoint.Shape;

/// <summary>
///     Handler for clearing shape format properties in PowerPoint presentations.
/// </summary>
[ResultType(typeof(SuccessResult))]
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
    public override object Execute(OperationContext<Presentation> context, OperationParameters parameters)
    {
        var p = ExtractClearPptShapeFormatParameters(parameters);
        var presentation = context.Document;
        var slide = PowerPointHelper.GetSlide(presentation, p.SlideIndex);

        PowerPointHelper.ValidateCollectionIndex(p.ShapeIndex, slide.Shapes.Count, "shapeIndex");

        var shape = slide.Shapes[p.ShapeIndex];

        if (p.ClearFill)
            shape.FillFormat.FillType = FillType.NoFill;

        if (p.ClearLine)
            shape.LineFormat.FillFormat.FillType = FillType.NoFill;

        MarkModified(context);

        return new SuccessResult { Message = $"Format cleared from shape {p.ShapeIndex} on slide {p.SlideIndex}." };
    }

    private static ClearPptShapeFormatParameters ExtractClearPptShapeFormatParameters(OperationParameters parameters)
    {
        return new ClearPptShapeFormatParameters(
            parameters.GetOptional("slideIndex", 0),
            parameters.GetRequired<int>("shapeIndex"),
            parameters.GetOptional("clearFill", true),
            parameters.GetOptional("clearLine", true));
    }

    private sealed record ClearPptShapeFormatParameters(int SlideIndex, int ShapeIndex, bool ClearFill, bool ClearLine);
}
