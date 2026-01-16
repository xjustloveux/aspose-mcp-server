using Aspose.Slides;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Helpers;

namespace AsposeMcpServer.Handlers.PowerPoint.Shape;

/// <summary>
///     Handler for getting detailed information about a specific shape.
/// </summary>
public class GetPptShapeDetailsHandler : OperationHandlerBase<Presentation>
{
    /// <inheritdoc />
    public override string Operation => "get_shape_details";

    /// <summary>
    ///     Gets detailed information about a specific shape.
    /// </summary>
    /// <param name="context">The presentation context.</param>
    /// <param name="parameters">
    ///     Required: shapeIndex.
    ///     Optional: slideIndex (default: 0).
    /// </param>
    /// <returns>JSON result with detailed shape information.</returns>
    public override string Execute(OperationContext<Presentation> context, OperationParameters parameters)
    {
        var p = ExtractGetPptShapeDetailsParameters(parameters);

        var presentation = context.Document;
        var slide = PowerPointHelper.GetSlide(presentation, p.SlideIndex);

        PowerPointHelper.ValidateCollectionIndex(p.ShapeIndex, slide.Shapes.Count, "shapeIndex");

        var shape = slide.Shapes[p.ShapeIndex];

        var result = new Dictionary<string, object?>
        {
            ["index"] = p.ShapeIndex,
            ["name"] = shape.Name,
            ["type"] = shape.GetType().Name,
            ["x"] = shape.X,
            ["y"] = shape.Y,
            ["width"] = shape.Width,
            ["height"] = shape.Height,
            ["rotation"] = shape.Rotation,
            ["hidden"] = shape.Hidden
        };

        if (shape is IAutoShape autoShape)
        {
            result["shapeType"] = autoShape.ShapeType.ToString();
            result["text"] = autoShape.TextFrame?.Text;
            result["fillType"] = autoShape.FillFormat?.FillType.ToString();
        }

        if (shape is ITable table)
        {
            result["rows"] = table.Rows.Count;
            result["columns"] = table.Columns.Count;
        }

        if (shape is IGroupShape groupShape) result["shapeCount"] = groupShape.Shapes.Count;

        return JsonResult(result);
    }

    /// <summary>
    ///     Extracts parameters for get shape details operation.
    /// </summary>
    /// <param name="parameters">The operation parameters.</param>
    /// <returns>The extracted parameters.</returns>
    private static GetPptShapeDetailsParameters ExtractGetPptShapeDetailsParameters(OperationParameters parameters)
    {
        return new GetPptShapeDetailsParameters(
            parameters.GetOptional("slideIndex", 0),
            parameters.GetRequired<int>("shapeIndex"));
    }

    /// <summary>
    ///     Parameters for get shape details operation.
    /// </summary>
    /// <param name="SlideIndex">The slide index (0-based).</param>
    /// <param name="ShapeIndex">The shape index.</param>
    private sealed record GetPptShapeDetailsParameters(int SlideIndex, int ShapeIndex);
}
