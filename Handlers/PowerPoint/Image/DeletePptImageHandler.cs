using Aspose.Slides;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Helpers;

namespace AsposeMcpServer.Handlers.PowerPoint.Image;

/// <summary>
///     Handler for deleting images from PowerPoint slides.
/// </summary>
public class DeletePptImageHandler : OperationHandlerBase<Presentation>
{
    /// <inheritdoc />
    public override string Operation => "delete";

    /// <summary>
    ///     Deletes an image from a slide.
    /// </summary>
    /// <param name="context">The presentation context.</param>
    /// <param name="parameters">
    ///     Required: slideIndex, imageIndex
    /// </param>
    /// <returns>Success message with deletion details.</returns>
    public override string Execute(OperationContext<Presentation> context, OperationParameters parameters)
    {
        var slideIndex = parameters.GetRequired<int>("slideIndex");
        var imageIndex = parameters.GetRequired<int>("imageIndex");

        var presentation = context.Document;
        var slide = PowerPointHelper.GetSlide(presentation, slideIndex);
        var pictures = PptImageHelper.GetPictureFrames(slide);

        PptImageHelper.ValidateImageIndex(imageIndex, slideIndex, pictures.Count);

        var pictureFrame = pictures[imageIndex];
        slide.Shapes.Remove(pictureFrame);

        MarkModified(context);

        return Success($"Image {imageIndex} deleted from slide {slideIndex}.");
    }
}
