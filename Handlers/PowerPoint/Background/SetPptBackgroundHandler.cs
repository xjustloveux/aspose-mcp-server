using System.Drawing;
using Aspose.Slides;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Helpers;

namespace AsposeMcpServer.Handlers.PowerPoint.Background;

/// <summary>
///     Handler for setting PowerPoint slide backgrounds.
/// </summary>
public class SetPptBackgroundHandler : OperationHandlerBase<Presentation>
{
    /// <inheritdoc />
    public override string Operation => "set";

    /// <summary>
    ///     Sets slide background with color or image.
    /// </summary>
    /// <param name="context">The presentation context.</param>
    /// <param name="parameters">
    ///     Optional: slideIndex (default: 0), color, imagePath, applyToAll (default: false)
    /// </param>
    /// <returns>Success message with background details.</returns>
    public override string Execute(OperationContext<Presentation> context, OperationParameters parameters)
    {
        var slideIndex = parameters.GetOptional("slideIndex", 0);
        var colorHex = parameters.GetOptional<string?>("color");
        var imagePath = parameters.GetOptional<string?>("imagePath");
        var applyToAll = parameters.GetOptional("applyToAll", false);

        if (string.IsNullOrWhiteSpace(colorHex) && string.IsNullOrWhiteSpace(imagePath))
            throw new ArgumentException("Please provide at least one of color or imagePath");

        var presentation = context.Document;

        IPPImage? img = null;
        if (!string.IsNullOrWhiteSpace(imagePath))
            img = presentation.Images.AddImage(File.ReadAllBytes(imagePath));

        Color? color = null;
        if (!string.IsNullOrWhiteSpace(colorHex))
            color = ColorHelper.ParseColor(colorHex);

        var slidesToUpdate = applyToAll
            ? presentation.Slides.ToList()
            : [PowerPointHelper.GetSlide(presentation, slideIndex)];

        foreach (var slide in slidesToUpdate)
            ApplyBackground(slide, color, img);

        MarkModified(context);

        var message = applyToAll
            ? $"Background applied to all {slidesToUpdate.Count} slides"
            : $"Background updated for slide {slideIndex}";

        return Success(message + ".");
    }

    /// <summary>
    ///     Applies background color or image to a slide.
    /// </summary>
    /// <param name="slide">The slide to apply background to.</param>
    /// <param name="color">The background color.</param>
    /// <param name="image">The background image.</param>
    private static void ApplyBackground(ISlide slide, Color? color, IPPImage? image)
    {
        slide.Background.Type = BackgroundType.OwnBackground;
        var fillFormat = slide.Background.FillFormat;

        if (image != null)
        {
            fillFormat.FillType = FillType.Picture;
            fillFormat.PictureFillFormat.PictureFillMode = PictureFillMode.Stretch;
            fillFormat.PictureFillFormat.Picture.Image = image;
        }
        else if (color.HasValue)
        {
            fillFormat.FillType = FillType.Solid;
            fillFormat.SolidFillColor.Color = color.Value;
        }
    }
}
