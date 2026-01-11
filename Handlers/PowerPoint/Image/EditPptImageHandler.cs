using Aspose.Slides;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Helpers;

namespace AsposeMcpServer.Handlers.PowerPoint.Image;

/// <summary>
///     Handler for editing images in PowerPoint slides.
/// </summary>
public class EditPptImageHandler : OperationHandlerBase<Presentation>
{
    /// <inheritdoc />
    public override string Operation => "edit";

    /// <summary>
    ///     Edits image properties with optional compression and resize.
    /// </summary>
    /// <param name="context">The presentation context.</param>
    /// <param name="parameters">
    ///     Required: slideIndex, imageIndex
    ///     Optional: imagePath, x, y, width, height, jpegQuality, maxWidth, maxHeight
    /// </param>
    /// <returns>Success message with edit details.</returns>
    public override string Execute(OperationContext<Presentation> context, OperationParameters parameters)
    {
        var slideIndex = parameters.GetRequired<int>("slideIndex");
        var imageIndex = parameters.GetRequired<int>("imageIndex");
        var imagePath = parameters.GetOptional<string?>("imagePath");
        var x = parameters.GetOptional("x", 100f);
        var y = parameters.GetOptional("y", 100f);
        var width = parameters.GetOptional<float?>("width");
        var height = parameters.GetOptional<float?>("height");
        var jpegQuality = parameters.GetOptional<int?>("jpegQuality");
        var maxWidth = parameters.GetOptional<int?>("maxWidth");
        var maxHeight = parameters.GetOptional<int?>("maxHeight");

        var presentation = context.Document;
        var slide = PowerPointHelper.GetSlide(presentation, slideIndex);
        var pictures = PptImageHelper.GetPictureFrames(slide);

        PptImageHelper.ValidateImageIndex(imageIndex, slideIndex, pictures.Count);

        var pictureFrame = pictures[imageIndex];
        List<string> changes = [];

        if (!string.IsNullOrEmpty(imagePath) && File.Exists(imagePath))
        {
            var newImage = PptImageHelper.ProcessAndAddImage(presentation, imagePath, jpegQuality, maxWidth, maxHeight,
                out var processingDetails);
            pictureFrame.PictureFormat.Picture.Image = newImage;
            changes.AddRange(processingDetails);
        }

        pictureFrame.X = x;
        pictureFrame.Y = y;

        if (width.HasValue)
        {
            pictureFrame.Width = width.Value;
            changes.Add($"width={width.Value}");
        }

        if (height.HasValue)
        {
            pictureFrame.Height = height.Value;
            changes.Add($"height={height.Value}");
        }

        MarkModified(context);

        var changesStr = changes.Count > 0 ? string.Join(", ", changes) : "position updated";
        return Success($"Image {imageIndex} on slide {slideIndex} updated ({changesStr}).");
    }
}
