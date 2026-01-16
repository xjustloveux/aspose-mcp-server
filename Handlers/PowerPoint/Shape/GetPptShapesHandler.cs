using Aspose.Slides;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Helpers;

namespace AsposeMcpServer.Handlers.PowerPoint.Shape;

/// <summary>
///     Handler for getting all shapes from a PowerPoint slide.
/// </summary>
public class GetPptShapesHandler : OperationHandlerBase<Presentation>
{
    /// <inheritdoc />
    public override string Operation => "get_shapes";

    /// <summary>
    ///     Gets all shapes from a slide.
    /// </summary>
    /// <param name="context">The presentation context.</param>
    /// <param name="parameters">
    ///     Optional: slideIndex (default: 0).
    /// </param>
    /// <returns>JSON result with shape information.</returns>
    public override string Execute(OperationContext<Presentation> context, OperationParameters parameters)
    {
        var p = ExtractGetPptShapesParameters(parameters);

        var presentation = context.Document;
        var slide = PowerPointHelper.GetSlide(presentation, p.SlideIndex);

        var shapes = new List<object>();
        for (var i = 0; i < slide.Shapes.Count; i++)
        {
            var shape = slide.Shapes[i];
            shapes.Add(new
            {
                index = i,
                name = shape.Name,
                type = shape.GetType().Name,
                x = shape.X,
                y = shape.Y,
                width = shape.Width,
                height = shape.Height,
                rotation = shape.Rotation
            });
        }

        var result = new
        {
            slideIndex = p.SlideIndex,
            count = shapes.Count,
            shapes
        };

        return JsonResult(result);
    }

    /// <summary>
    ///     Extracts parameters for get shapes operation.
    /// </summary>
    /// <param name="parameters">The operation parameters.</param>
    /// <returns>The extracted parameters.</returns>
    private static GetPptShapesParameters ExtractGetPptShapesParameters(OperationParameters parameters)
    {
        return new GetPptShapesParameters(parameters.GetOptional("slideIndex", 0));
    }

    /// <summary>
    ///     Parameters for get shapes operation.
    /// </summary>
    /// <param name="SlideIndex">The slide index (0-based).</param>
    private sealed record GetPptShapesParameters(int SlideIndex);
}
