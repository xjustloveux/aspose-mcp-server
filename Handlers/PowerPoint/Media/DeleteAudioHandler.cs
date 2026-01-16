using Aspose.Slides;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Helpers;

namespace AsposeMcpServer.Handlers.PowerPoint.Media;

/// <summary>
///     Handler for deleting audio from PowerPoint presentations.
/// </summary>
public class DeleteAudioHandler : OperationHandlerBase<Presentation>
{
    /// <inheritdoc />
    public override string Operation => "delete_audio";

    /// <summary>
    ///     Deletes an audio frame from a slide.
    /// </summary>
    /// <param name="context">The presentation context.</param>
    /// <param name="parameters">
    ///     Required: shapeIndex.
    ///     Optional: slideIndex (default: 0).
    /// </param>
    /// <returns>Success message with deletion details.</returns>
    public override string Execute(OperationContext<Presentation> context, OperationParameters parameters)
    {
        var p = ExtractDeleteAudioParameters(parameters);

        var presentation = context.Document;
        var slide = PowerPointHelper.GetSlide(presentation, p.SlideIndex);

        PowerPointHelper.ValidateCollectionIndex(p.ShapeIndex, slide.Shapes.Count, "shapeIndex");

        if (slide.Shapes[p.ShapeIndex] is not IAudioFrame)
            throw new ArgumentException($"Shape at index {p.ShapeIndex} is not an audio frame");

        slide.Shapes.RemoveAt(p.ShapeIndex);

        MarkModified(context);

        return Success($"Audio deleted from slide {p.SlideIndex}.");
    }

    /// <summary>
    ///     Extracts delete audio parameters from operation parameters.
    /// </summary>
    /// <param name="parameters">The operation parameters.</param>
    /// <returns>The extracted delete audio parameters.</returns>
    private static DeleteAudioParameters ExtractDeleteAudioParameters(OperationParameters parameters)
    {
        return new DeleteAudioParameters(
            parameters.GetOptional("slideIndex", 0),
            parameters.GetRequired<int>("shapeIndex"));
    }

    /// <summary>
    ///     Record for holding delete audio parameters.
    /// </summary>
    /// <param name="SlideIndex">The slide index.</param>
    /// <param name="ShapeIndex">The shape index.</param>
    private sealed record DeleteAudioParameters(int SlideIndex, int ShapeIndex);
}
