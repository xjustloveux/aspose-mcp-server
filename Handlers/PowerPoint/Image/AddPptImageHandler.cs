using Aspose.Slides;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Helpers;

namespace AsposeMcpServer.Handlers.PowerPoint.Image;

/// <summary>
///     Handler for adding images to PowerPoint slides.
/// </summary>
public class AddPptImageHandler : OperationHandlerBase<Presentation>
{
    /// <inheritdoc />
    public override string Operation => "add";

    /// <summary>
    ///     Adds an image to a slide.
    /// </summary>
    /// <param name="context">The presentation context.</param>
    /// <param name="parameters">
    ///     Required: slideIndex, imagePath
    ///     Optional: x, y, width, height
    /// </param>
    /// <returns>Success message with image addition details.</returns>
    public override string Execute(OperationContext<Presentation> context, OperationParameters parameters)
    {
        var slideIndex = parameters.GetRequired<int>("slideIndex");
        var imagePath = parameters.GetRequired<string>("imagePath");
        var x = parameters.GetOptional("x", 100f);
        var y = parameters.GetOptional("y", 100f);
        var width = parameters.GetOptional<float?>("width");
        var height = parameters.GetOptional<float?>("height");

        if (!File.Exists(imagePath))
            throw new FileNotFoundException($"Image file not found: {imagePath}");

        var presentation = context.Document;
        var slide = PowerPointHelper.GetSlide(presentation, slideIndex);

        IPPImage pictureImage;
        int pixelWidth, pixelHeight;

        using (var fileStream = new FileStream(imagePath, FileMode.Open, FileAccess.Read))
        {
            pictureImage = presentation.Images.AddImage(fileStream);
            pixelWidth = pictureImage.Width;
            pixelHeight = pictureImage.Height;
        }

        var (finalWidth, finalHeight) = PptImageHelper.CalculateDimensions(width, height, pixelWidth, pixelHeight);

        slide.Shapes.AddPictureFrame(ShapeType.Rectangle, x, y, finalWidth, finalHeight, pictureImage);

        MarkModified(context);

        return Success($"Image added to slide {slideIndex}.");
    }
}
