using Aspose.Slides;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.ShapeDetailProviders;
using AsposeMcpServer.Helpers.PowerPoint;
using AsposeMcpServer.Results.PowerPoint.Shape;

namespace AsposeMcpServer.Handlers.PowerPoint.Shape;

/// <summary>
///     Handler for getting detailed information about a specific shape.
/// </summary>
[ResultType(typeof(GetShapeDetailsResult))]
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
    public override object Execute(OperationContext<Presentation> context, OperationParameters parameters)
    {
        var p = ExtractGetPptShapeDetailsParameters(parameters);

        var presentation = context.Document;
        var slide = PowerPointHelper.GetSlide(presentation, p.SlideIndex);

        PowerPointHelper.ValidateCollectionIndex(p.ShapeIndex, slide.Shapes.Count, "shapeIndex");

        var shape = slide.Shapes[p.ShapeIndex];

        var shapeInfo = ShapeDetailProviderFactory.GetShapeDetails(shape, presentation);

        return new GetShapeDetailsResult
        {
            Index = p.ShapeIndex,
            Name = shape.Name,
            Type = shapeInfo.TypeName,
            X = shape.X,
            Y = shape.Y,
            Width = shape.Width,
            Height = shape.Height,
            Rotation = shape.Rotation,
            Hidden = shape.Hidden,
            AlternativeText = string.IsNullOrEmpty(shape.AlternativeText) ? null : shape.AlternativeText,
            FlipHorizontal = MapNullableBool(shape.Frame.FlipH),
            FlipVertical = MapNullableBool(shape.Frame.FlipV),
            Details = shapeInfo.Details
        };
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
    ///     Maps a <see cref="NullableBool" /> value to a nullable boolean.
    /// </summary>
    /// <param name="value">The NullableBool value.</param>
    /// <returns>true, false, or null if not defined.</returns>
    private static bool? MapNullableBool(NullableBool value)
    {
        return value switch
        {
            NullableBool.True => true,
            NullableBool.False => false,
            _ => null
        };
    }

    /// <summary>
    ///     Parameters for get shape details operation.
    /// </summary>
    /// <param name="SlideIndex">The slide index (0-based).</param>
    /// <param name="ShapeIndex">The shape index.</param>
    private sealed record GetPptShapeDetailsParameters(int SlideIndex, int ShapeIndex);
}
