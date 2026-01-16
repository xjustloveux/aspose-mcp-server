using Aspose.Slides;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Helpers;

namespace AsposeMcpServer.Handlers.PowerPoint.Shape;

/// <summary>
///     Handler for aligning multiple shapes.
/// </summary>
public class AlignPptShapesHandler : OperationHandlerBase<Presentation>
{
    /// <inheritdoc />
    public override string Operation => "align";

    /// <summary>
    ///     Aligns multiple shapes.
    /// </summary>
    /// <param name="context">The presentation context.</param>
    /// <param name="parameters">
    ///     Required: slideIndex, shapeIndices (at least 2 shapes), align
    ///     Optional: alignToSlide (default: false)
    /// </param>
    /// <returns>Success message with alignment details.</returns>
    public override string Execute(OperationContext<Presentation> context, OperationParameters parameters)
    {
        var p = ExtractAlignPptShapesParameters(parameters);

        if (p.ShapeIndices.Length < 2)
            throw new ArgumentException("At least 2 shapes are required for alignment");

        var presentation = context.Document;
        var slide = PowerPointHelper.GetSlide(presentation, p.SlideIndex);

        foreach (var idx in p.ShapeIndices)
            PowerPointHelper.ValidateCollectionIndex(idx, slide.Shapes.Count, "shapeIndex");

        var shapes = p.ShapeIndices.Select(idx => slide.Shapes[idx]).ToArray();

        var refBox = p.AlignToSlide
            ? new { X = 0f, Y = 0f, W = presentation.SlideSize.Size.Width, H = presentation.SlideSize.Size.Height }
            : new
            {
                X = shapes.Min(s => s.X),
                Y = shapes.Min(s => s.Y),
                W = shapes.Max(s => s.X + s.Width) - shapes.Min(s => s.X),
                H = shapes.Max(s => s.Y + s.Height) - shapes.Min(s => s.Y)
            };

        foreach (var s in shapes)
            switch (p.Align.ToLower())
            {
                case "left": s.X = refBox.X; break;
                case "center": s.X = refBox.X + (refBox.W - s.Width) / 2f; break;
                case "right": s.X = refBox.X + refBox.W - s.Width; break;
                case "top": s.Y = refBox.Y; break;
                case "middle": s.Y = refBox.Y + (refBox.H - s.Height) / 2f; break;
                case "bottom": s.Y = refBox.Y + refBox.H - s.Height; break;
                default:
                    throw new ArgumentException("align must be: left|center|right|top|middle|bottom");
            }

        MarkModified(context);

        return Success($"Aligned {p.ShapeIndices.Length} shapes: {p.Align}.");
    }

    private static AlignPptShapesParameters ExtractAlignPptShapesParameters(OperationParameters parameters)
    {
        return new AlignPptShapesParameters(
            parameters.GetRequired<int>("slideIndex"),
            parameters.GetRequired<int[]>("shapeIndices"),
            parameters.GetRequired<string>("align"),
            parameters.GetOptional("alignToSlide", false));
    }

    private sealed record AlignPptShapesParameters(int SlideIndex, int[] ShapeIndices, string Align, bool AlignToSlide);
}
