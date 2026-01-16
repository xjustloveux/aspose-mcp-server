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
        var p = ExtractEditParameters(parameters);

        var presentation = context.Document;
        var slide = PowerPointHelper.GetSlide(presentation, p.SlideIndex);
        var pictures = PptImageHelper.GetPictureFrames(slide);

        PptImageHelper.ValidateImageIndex(p.ImageIndex, p.SlideIndex, pictures.Count);

        var pictureFrame = pictures[p.ImageIndex];
        List<string> changes = [];

        if (!string.IsNullOrEmpty(p.ImagePath) && File.Exists(p.ImagePath))
        {
            var newImage = PptImageHelper.ProcessAndAddImage(presentation, p.ImagePath, p.JpegQuality, p.MaxWidth,
                p.MaxHeight,
                out var processingDetails);
            pictureFrame.PictureFormat.Picture.Image = newImage;
            changes.AddRange(processingDetails);
        }

        pictureFrame.X = p.X;
        pictureFrame.Y = p.Y;

        if (p.Width.HasValue)
        {
            pictureFrame.Width = p.Width.Value;
            changes.Add($"width={p.Width.Value}");
        }

        if (p.Height.HasValue)
        {
            pictureFrame.Height = p.Height.Value;
            changes.Add($"height={p.Height.Value}");
        }

        MarkModified(context);

        var changesStr = changes.Count > 0 ? string.Join(", ", changes) : "position updated";
        return Success($"Image {p.ImageIndex} on slide {p.SlideIndex} updated ({changesStr}).");
    }

    /// <summary>
    ///     Extracts edit parameters from operation parameters.
    /// </summary>
    /// <param name="parameters">The operation parameters.</param>
    /// <returns>The extracted edit parameters.</returns>
    private static EditParameters ExtractEditParameters(OperationParameters parameters)
    {
        return new EditParameters(
            parameters.GetRequired<int>("slideIndex"),
            parameters.GetRequired<int>("imageIndex"),
            parameters.GetOptional<string?>("imagePath"),
            parameters.GetOptional("x", 100f),
            parameters.GetOptional("y", 100f),
            parameters.GetOptional<float?>("width"),
            parameters.GetOptional<float?>("height"),
            parameters.GetOptional<int?>("jpegQuality"),
            parameters.GetOptional<int?>("maxWidth"),
            parameters.GetOptional<int?>("maxHeight")
        );
    }

    /// <summary>
    ///     Record for holding edit image parameters.
    /// </summary>
    /// <param name="SlideIndex">The slide index.</param>
    /// <param name="ImageIndex">The image index.</param>
    /// <param name="ImagePath">The optional new image path.</param>
    /// <param name="X">The X coordinate.</param>
    /// <param name="Y">The Y coordinate.</param>
    /// <param name="Width">The optional width.</param>
    /// <param name="Height">The optional height.</param>
    /// <param name="JpegQuality">The optional JPEG quality.</param>
    /// <param name="MaxWidth">The optional maximum width.</param>
    /// <param name="MaxHeight">The optional maximum height.</param>
    private record EditParameters(
        int SlideIndex,
        int ImageIndex,
        string? ImagePath,
        float X,
        float Y,
        float? Width,
        float? Height,
        int? JpegQuality,
        int? MaxWidth,
        int? MaxHeight);
}
