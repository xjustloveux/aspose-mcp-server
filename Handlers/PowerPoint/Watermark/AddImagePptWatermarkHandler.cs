using Aspose.Slides;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Helpers;
using AsposeMcpServer.Results.Common;

namespace AsposeMcpServer.Handlers.PowerPoint.Watermark;

/// <summary>
///     Handler for adding an image watermark to a PowerPoint presentation.
/// </summary>
[ResultType(typeof(SuccessResult))]
public class AddImagePptWatermarkHandler : OperationHandlerBase<Presentation>
{
    /// <inheritdoc />
    public override string Operation => "add_image";

    /// <summary>
    ///     Adds an image watermark to all slides.
    /// </summary>
    /// <param name="context">The presentation context.</param>
    /// <param name="parameters">
    ///     Required: imagePath.
    ///     Optional: width (default: 200), height (default: 200).
    /// </param>
    /// <returns>Success message with watermark details.</returns>
    /// <exception cref="ArgumentException">Thrown when the image file does not exist.</exception>
    public override object Execute(OperationContext<Presentation> context, OperationParameters parameters)
    {
        var imagePath = parameters.GetRequired<string>("imagePath");
        var width = parameters.GetOptional("width", 200f);
        var height = parameters.GetOptional("height", 200f);

        SecurityHelper.ValidateFilePath(imagePath, "imagePath");

        if (!File.Exists(imagePath))
            throw new ArgumentException($"Image file not found: {imagePath}");

        var presentation = context.Document;
        var imageBytes = File.ReadAllBytes(imagePath);
        var ppImage = presentation.Images.AddImage(imageBytes);

        var slideCount = 0;
        foreach (var slide in presentation.Slides)
        {
            AddImageWatermarkToSlide(slide, ppImage, width, height);
            slideCount++;
        }

        MarkModified(context);
        return new SuccessResult { Message = $"Image watermark added to {slideCount} slide(s)." };
    }

    /// <summary>
    ///     Adds an image watermark to a single slide.
    /// </summary>
    /// <param name="slide">The slide to add the watermark to.</param>
    /// <param name="image">The image to use as watermark.</param>
    /// <param name="width">The width of the watermark in points.</param>
    /// <param name="height">The height of the watermark in points.</param>
    private static void AddImageWatermarkToSlide(ISlide slide, IPPImage image, float width, float height)
    {
        var slideWidth = slide.Presentation.SlideSize.Size.Width;
        var slideHeight = slide.Presentation.SlideSize.Size.Height;

        var x = (slideWidth - width) / 2f;
        var y = (slideHeight - height) / 2f;

        var pictureFrame = slide.Shapes.AddPictureFrame(ShapeType.Rectangle, x, y, width, height, image);
        pictureFrame.Name = $"{AddTextPptWatermarkHandler.WatermarkPrefix}IMAGE_{Guid.NewGuid():N}";

        pictureFrame.ShapeLock.SelectLocked = true;
        pictureFrame.ShapeLock.SizeLocked = true;
        pictureFrame.ShapeLock.PositionLocked = true;
    }
}
