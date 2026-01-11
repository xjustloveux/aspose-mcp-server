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
        var slideIndex = parameters.GetOptional("slideIndex", 0);

        var presentation = context.Document;
        var slide = PowerPointHelper.GetSlide(presentation, slideIndex);

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
            slideIndex,
            count = shapes.Count,
            shapes
        };

        return JsonResult(result);
    }
}
