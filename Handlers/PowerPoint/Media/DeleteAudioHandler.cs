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
        var slideIndex = parameters.GetOptional("slideIndex", 0);
        var shapeIndex = parameters.GetRequired<int>("shapeIndex");

        var presentation = context.Document;
        var slide = PowerPointHelper.GetSlide(presentation, slideIndex);

        PowerPointHelper.ValidateCollectionIndex(shapeIndex, slide.Shapes.Count, "shapeIndex");

        if (slide.Shapes[shapeIndex] is not IAudioFrame)
            throw new ArgumentException($"Shape at index {shapeIndex} is not an audio frame");

        slide.Shapes.RemoveAt(shapeIndex);

        MarkModified(context);

        return Success($"Audio deleted from slide {slideIndex}.");
    }
}
