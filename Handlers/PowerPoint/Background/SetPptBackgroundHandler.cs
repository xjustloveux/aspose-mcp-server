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
        var p = ExtractSetBackgroundParameters(parameters);

        if (string.IsNullOrWhiteSpace(p.Color) && string.IsNullOrWhiteSpace(p.ImagePath))
            throw new ArgumentException("Please provide at least one of color or imagePath");

        var presentation = context.Document;

        IPPImage? img = null;
        if (!string.IsNullOrWhiteSpace(p.ImagePath))
            img = presentation.Images.AddImage(File.ReadAllBytes(p.ImagePath));

        Color? color = null;
        if (!string.IsNullOrWhiteSpace(p.Color))
            color = ColorHelper.ParseColor(p.Color);

        var slidesToUpdate = p.ApplyToAll
            ? presentation.Slides.ToList()
            : [PowerPointHelper.GetSlide(presentation, p.SlideIndex)];

        foreach (var slide in slidesToUpdate)
            ApplyBackground(slide, color, img);

        MarkModified(context);

        var message = p.ApplyToAll
            ? $"Background applied to all {slidesToUpdate.Count} slides"
            : $"Background updated for slide {p.SlideIndex}";

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

    /// <summary>
    ///     Extracts set background parameters from operation parameters.
    /// </summary>
    /// <param name="parameters">The operation parameters.</param>
    /// <returns>The extracted set background parameters.</returns>
    private static SetBackgroundParameters ExtractSetBackgroundParameters(OperationParameters parameters)
    {
        return new SetBackgroundParameters(
            parameters.GetOptional("slideIndex", 0),
            parameters.GetOptional<string?>("color"),
            parameters.GetOptional<string?>("imagePath"),
            parameters.GetOptional("applyToAll", false));
    }

    /// <summary>
    ///     Record for holding set background parameters.
    /// </summary>
    /// <param name="SlideIndex">The slide index.</param>
    /// <param name="Color">The optional background color.</param>
    /// <param name="ImagePath">The optional background image path.</param>
    /// <param name="ApplyToAll">Whether to apply to all slides.</param>
    private sealed record SetBackgroundParameters(int SlideIndex, string? Color, string? ImagePath, bool ApplyToAll);
}
