using Aspose.Slides;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Helpers.PowerPoint;
using AsposeMcpServer.Results.PowerPoint.Shape;

namespace AsposeMcpServer.Handlers.PowerPoint.Shape;

/// <summary>
///     Handler for getting all shapes from a PowerPoint slide.
/// </summary>
[ResultType(typeof(GetShapesResult))]
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
    public override object Execute(OperationContext<Presentation> context, OperationParameters parameters)
    {
        var p = ExtractGetPptShapesParameters(parameters);

        var presentation = context.Document;
        var slide = PowerPointHelper.GetSlide(presentation, p.SlideIndex);

        var shapes = new List<GetShapeInfo>();
        for (var i = 0; i < slide.Shapes.Count; i++)
        {
            var shape = slide.Shapes[i];
            shapes.Add(new GetShapeInfo
            {
                Index = i,
                Name = shape.Name,
                Type = shape.GetType().Name,
                X = shape.X,
                Y = shape.Y,
                Width = shape.Width,
                Height = shape.Height,
                Rotation = shape.Rotation
            });
        }

        return new GetShapesResult
        {
            SlideIndex = p.SlideIndex,
            Count = shapes.Count,
            Shapes = shapes
        };
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
