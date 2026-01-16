using Aspose.Slides;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Helpers;

namespace AsposeMcpServer.Handlers.PowerPoint.Media;

/// <summary>
///     Handler for deleting video from PowerPoint presentations.
/// </summary>
public class DeleteVideoHandler : OperationHandlerBase<Presentation>
{
    /// <inheritdoc />
    public override string Operation => "delete_video";

    /// <summary>
    ///     Deletes a video frame from a slide.
    /// </summary>
    /// <param name="context">The presentation context.</param>
    /// <param name="parameters">
    ///     Required: shapeIndex.
    ///     Optional: slideIndex (default: 0).
    /// </param>
    /// <returns>Success message with deletion details.</returns>
    public override string Execute(OperationContext<Presentation> context, OperationParameters parameters)
    {
        var p = ExtractDeleteVideoParameters(parameters);

        var presentation = context.Document;
        var slide = PowerPointHelper.GetSlide(presentation, p.SlideIndex);

        PowerPointHelper.ValidateCollectionIndex(p.ShapeIndex, slide.Shapes.Count, "shapeIndex");

        if (slide.Shapes[p.ShapeIndex] is not IVideoFrame)
            throw new ArgumentException($"Shape at index {p.ShapeIndex} is not a video frame");

        slide.Shapes.RemoveAt(p.ShapeIndex);

        MarkModified(context);

        return Success($"Video deleted from slide {p.SlideIndex}.");
    }

    /// <summary>
    ///     Extracts delete video parameters from operation parameters.
    /// </summary>
    /// <param name="parameters">The operation parameters.</param>
    /// <returns>The extracted delete video parameters.</returns>
    private static DeleteVideoParameters ExtractDeleteVideoParameters(OperationParameters parameters)
    {
        return new DeleteVideoParameters(
            parameters.GetOptional("slideIndex", 0),
            parameters.GetRequired<int>("shapeIndex"));
    }

    /// <summary>
    ///     Record for holding delete video parameters.
    /// </summary>
    /// <param name="SlideIndex">The slide index.</param>
    /// <param name="ShapeIndex">The shape index.</param>
    private sealed record DeleteVideoParameters(int SlideIndex, int ShapeIndex);
}
