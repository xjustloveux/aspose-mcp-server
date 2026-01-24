using System.Drawing;
using System.Drawing.Imaging;
using System.Security.Cryptography;
using Aspose.Slides;

namespace AsposeMcpServer.Helpers.PowerPoint;

/// <summary>
///     Helper class for common PowerPoint image operations to reduce code duplication.
/// </summary>
public static class PptImageHelper
{
    /// <summary>
    ///     Calculates final dimensions maintaining aspect ratio.
    /// </summary>
    /// <param name="width">The requested width (optional).</param>
    /// <param name="height">The requested height (optional).</param>
    /// <param name="pixelWidth">The original image width in pixels.</param>
    /// <param name="pixelHeight">The original image height in pixels.</param>
    /// <returns>A tuple containing the calculated width and height.</returns>
    public static (float width, float height) CalculateDimensions(float? width, float? height, int pixelWidth,
        int pixelHeight)
    {
        if (width.HasValue && height.HasValue)
            return (width.Value, height.Value);

        if (width.HasValue)
        {
            var ratio = pixelWidth > 0 ? (float)pixelHeight / pixelWidth : 1;
            return (width.Value, width.Value * ratio);
        }

        if (height.HasValue)
        {
            var ratio = pixelHeight > 0 ? (float)pixelWidth / pixelHeight : 1;
            return (height.Value * ratio, height.Value);
        }

        var defaultWidth = 300f;
        var defaultRatio = pixelWidth > 0 ? (float)pixelHeight / pixelWidth : 1;
        return (defaultWidth, defaultWidth * defaultRatio);
    }

    /// <summary>
    ///     Calculates new image size maintaining aspect ratio within max bounds.
    /// </summary>
    /// <param name="width">The original width.</param>
    /// <param name="height">The original height.</param>
    /// <param name="maxWidth">The maximum width constraint (optional).</param>
    /// <param name="maxHeight">The maximum height constraint (optional).</param>
    /// <returns>The calculated size within the constraints.</returns>
    public static Size CalculateResizeSize(int width, int height, int? maxWidth, int? maxHeight)
    {
        var newWidth = (double)width;
        var newHeight = (double)height;

        if (maxWidth.HasValue && width > maxWidth.Value)
        {
            var ratio = (double)maxWidth.Value / width;
            newWidth = maxWidth.Value;
            newHeight *= ratio;
        }

        if (maxHeight.HasValue && newHeight > maxHeight.Value)
        {
            var ratio = maxHeight.Value / newHeight;
            newHeight = maxHeight.Value;
            newWidth *= ratio;
        }

        return new Size((int)Math.Round(newWidth), (int)Math.Round(newHeight));
    }

    /// <summary>
    ///     Processes image with optional compression/resize and adds to presentation.
    /// </summary>
    /// <param name="presentation">The presentation to add the image to.</param>
    /// <param name="imagePath">The path to the image file.</param>
    /// <param name="jpegQuality">The JPEG quality 10-100 (optional).</param>
    /// <param name="maxWidth">The maximum width in pixels for resize (optional).</param>
    /// <param name="maxHeight">The maximum height in pixels for resize (optional).</param>
    /// <param name="processingDetails">Output list of processing details performed.</param>
    /// <returns>The processed image added to the presentation.</returns>
    public static IPPImage ProcessAndAddImage(IPresentation presentation, string imagePath, int? jpegQuality,
        int? maxWidth, int? maxHeight, out List<string> processingDetails)
    {
        processingDetails = [];

        if (jpegQuality.HasValue || maxWidth.HasValue || maxHeight.HasValue)
        {
            // CA1416 - System.Drawing.Common is Windows-only, cross-platform support not required
#pragma warning disable CA1416
            using var fileStream = new FileStream(imagePath, FileMode.Open, FileAccess.Read);
            using var src = Image.FromStream(fileStream);

            var processedImage = src;
            var needsDispose = false;

            if (maxWidth.HasValue || maxHeight.HasValue)
            {
                var newSize = CalculateResizeSize(src.Width, src.Height, maxWidth, maxHeight);
                if (newSize.Width != src.Width || newSize.Height != src.Height)
                {
                    processedImage = new Bitmap(src, newSize);
                    needsDispose = true;
                    processingDetails.Add($"resized to {newSize.Width}x{newSize.Height}");
                }
            }

            using var ms = new MemoryStream();
            if (jpegQuality.HasValue)
            {
                var quality = Math.Clamp(jpegQuality.Value, 10, 100);
                var encoder = ImageCodecInfo.GetImageEncoders().First(c => c.FormatID == ImageFormat.Jpeg.Guid);
                var encParams = new EncoderParameters(1);
                encParams.Param[0] = new EncoderParameter(Encoder.Quality, quality);
                processedImage.Save(ms, encoder, encParams);
                processingDetails.Add($"quality={quality}");
            }
            else
            {
                processedImage.Save(ms, ImageFormat.Png);
            }

            if (needsDispose)
                processedImage.Dispose();
#pragma warning restore CA1416

            ms.Position = 0;
            processingDetails.Insert(0, "image replaced");
            return presentation.Images.AddImage(ms);
        }

        using var fs = new FileStream(imagePath, FileMode.Open, FileAccess.Read);
        processingDetails.Add("image replaced");
        return presentation.Images.AddImage(fs);
    }

    /// <summary>
    ///     Parses comma-separated slide indexes string.
    /// </summary>
    /// <param name="slideIndexesStr">Comma-separated slide indexes string.</param>
    /// <param name="totalSlides">The total number of slides in the presentation.</param>
    /// <returns>A list of valid slide indexes.</returns>
    /// <exception cref="ArgumentException">Thrown when a slide index is invalid or out of range.</exception>
    public static List<int> ParseSlideIndexes(string? slideIndexesStr, int totalSlides)
    {
        if (string.IsNullOrWhiteSpace(slideIndexesStr))
            return Enumerable.Range(0, totalSlides).ToList();

        List<int> indexes = [];
        foreach (var part in slideIndexesStr.Split(',', StringSplitOptions.RemoveEmptyEntries))
        {
            if (!int.TryParse(part.Trim(), out var index))
                throw new ArgumentException($"Invalid slide index: '{part}'");

            if (index < 0 || index >= totalSlides)
                throw new ArgumentException($"slideIndex {index} must be between 0 and {totalSlides - 1}");

            if (!indexes.Contains(index))
                indexes.Add(index);
        }

        return indexes;
    }

    /// <summary>
    ///     Computes MD5 hash of image binary data for duplicate detection.
    /// </summary>
    /// <param name="data">The image binary data.</param>
    /// <returns>The hexadecimal hash string.</returns>
    public static string ComputeImageHash(byte[] data)
    {
        var hashBytes = MD5.HashData(data);
        return Convert.ToHexString(hashBytes);
    }

    /// <summary>
    ///     Gets picture frames from a slide.
    /// </summary>
    /// <param name="slide">The slide to get picture frames from.</param>
    /// <returns>List of picture frames.</returns>
    public static List<PictureFrame> GetPictureFrames(ISlide slide)
    {
        return slide.Shapes.OfType<PictureFrame>().ToList();
    }

    /// <summary>
    ///     Validates image index and throws if out of range.
    /// </summary>
    /// <param name="imageIndex">The image index to validate.</param>
    /// <param name="slideIndex">The slide index for error message.</param>
    /// <param name="pictureCount">The number of pictures on the slide.</param>
    /// <exception cref="ArgumentException">Thrown when image index is out of range.</exception>
    public static void ValidateImageIndex(int imageIndex, int slideIndex, int pictureCount)
    {
        if (imageIndex < 0 || imageIndex >= pictureCount)
            throw new ArgumentException(
                $"imageIndex {imageIndex} is out of range. Slide {slideIndex} has {pictureCount} image(s).");
    }
}
