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
        var p = ExtractDeleteParameters(parameters);

        var presentation = context.Document;
        var slide = PowerPointHelper.GetSlide(presentation, p.SlideIndex);
        var pictures = PptImageHelper.GetPictureFrames(slide);

        PptImageHelper.ValidateImageIndex(p.ImageIndex, p.SlideIndex, pictures.Count);

        var pictureFrame = pictures[p.ImageIndex];
        slide.Shapes.Remove(pictureFrame);

        MarkModified(context);

        return Success($"Image {p.ImageIndex} deleted from slide {p.SlideIndex}.");
    }

    /// <summary>
    ///     Extracts delete parameters from operation parameters.
    /// </summary>
    /// <param name="parameters">The operation parameters.</param>
    /// <returns>The extracted delete parameters.</returns>
    private static DeleteParameters ExtractDeleteParameters(OperationParameters parameters)
    {
        return new DeleteParameters(
            parameters.GetRequired<int>("slideIndex"),
            parameters.GetRequired<int>("imageIndex")
        );
    }

    /// <summary>
    ///     Record for holding delete image parameters.
    /// </summary>
    /// <param name="SlideIndex">The slide index.</param>
    /// <param name="ImageIndex">The image index.</param>
    private record DeleteParameters(int SlideIndex, int ImageIndex);
}
