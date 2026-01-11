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
        var slideIndex = parameters.GetOptional("slideIndex", 0);
        var shapeIndex = parameters.GetRequired<int>("shapeIndex");

        var presentation = context.Document;
        var slide = PowerPointHelper.GetSlide(presentation, slideIndex);

        PowerPointHelper.ValidateCollectionIndex(shapeIndex, slide.Shapes.Count, "shapeIndex");

        var shape = slide.Shapes[shapeIndex];

        var result = new Dictionary<string, object?>
        {
            ["index"] = shapeIndex,
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
}
