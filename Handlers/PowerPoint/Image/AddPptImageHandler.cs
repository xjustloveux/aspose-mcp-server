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
        var p = ExtractAddParameters(parameters);

        if (!File.Exists(p.ImagePath))
            throw new FileNotFoundException($"Image file not found: {p.ImagePath}");

        var presentation = context.Document;
        var slide = PowerPointHelper.GetSlide(presentation, p.SlideIndex);

        IPPImage pictureImage;
        int pixelWidth, pixelHeight;

        using (var fileStream = new FileStream(p.ImagePath, FileMode.Open, FileAccess.Read))
        {
            pictureImage = presentation.Images.AddImage(fileStream);
            pixelWidth = pictureImage.Width;
            pixelHeight = pictureImage.Height;
        }

        var (finalWidth, finalHeight) = PptImageHelper.CalculateDimensions(p.Width, p.Height, pixelWidth, pixelHeight);

        slide.Shapes.AddPictureFrame(ShapeType.Rectangle, p.X, p.Y, finalWidth, finalHeight, pictureImage);

        MarkModified(context);

        return Success($"Image added to slide {p.SlideIndex}.");
    }

    /// <summary>
    ///     Extracts add parameters from operation parameters.
    /// </summary>
    /// <param name="parameters">The operation parameters.</param>
    /// <returns>The extracted add parameters.</returns>
    private static AddParameters ExtractAddParameters(OperationParameters parameters)
    {
        return new AddParameters(
            parameters.GetRequired<int>("slideIndex"),
            parameters.GetRequired<string>("imagePath"),
            parameters.GetOptional("x", 100f),
            parameters.GetOptional("y", 100f),
            parameters.GetOptional<float?>("width"),
            parameters.GetOptional<float?>("height")
        );
    }

    /// <summary>
    ///     Record for holding add image parameters.
    /// </summary>
    /// <param name="SlideIndex">The slide index.</param>
    /// <param name="ImagePath">The image file path.</param>
    /// <param name="X">The X coordinate.</param>
    /// <param name="Y">The Y coordinate.</param>
    /// <param name="Width">The optional width.</param>
    /// <param name="Height">The optional height.</param>
    private sealed record AddParameters(
        int SlideIndex,
        string ImagePath,
        float X,
        float Y,
        float? Width,
        float? Height);
}
