using Aspose.Cells.Drawing;

namespace AsposeMcpServer.Handlers.Excel.Image;

/// <summary>
///     Helper class for Excel image operations.
/// </summary>
public static class ExcelImageHelper
{
    /// <summary>
    ///     Set of supported image file extensions.
    /// </summary>
    public static readonly HashSet<string> SupportedImageExtensions = new(StringComparer.OrdinalIgnoreCase)
    {
        ".png", ".jpg", ".jpeg", ".gif", ".bmp", ".tiff", ".tif", ".emf", ".wmf"
    };

    /// <summary>
    ///     Mapping of file extensions to Aspose.Cells ImageType enum values.
    /// </summary>
    public static readonly Dictionary<string, ImageType> ExtensionToImageType = new(StringComparer.OrdinalIgnoreCase)
    {
        { ".png", ImageType.Png },
        { ".jpg", ImageType.Jpeg },
        { ".jpeg", ImageType.Jpeg },
        { ".gif", ImageType.Gif },
        { ".bmp", ImageType.Bmp },
        { ".tiff", ImageType.Tiff },
        { ".tif", ImageType.Tiff },
        { ".emf", ImageType.Emf },
        { ".wmf", ImageType.Wmf }
    };

    /// <summary>
    ///     Validates that the image file has a supported format.
    /// </summary>
    /// <param name="imagePath">The path to the image file to validate.</param>
    /// <exception cref="ArgumentException">Thrown when the image format is not supported.</exception>
    public static void ValidateImageFormat(string imagePath)
    {
        var extension = Path.GetExtension(imagePath);
        if (string.IsNullOrEmpty(extension) || !SupportedImageExtensions.Contains(extension))
            throw new ArgumentException(
                $"Unsupported image format: '{extension}'. Supported formats: {string.Join(", ", SupportedImageExtensions)}");
    }

    /// <summary>
    ///     Validates that the image index is within the valid range.
    /// </summary>
    /// <param name="imageIndex">The image index to validate.</param>
    /// <param name="picturesCount">The total number of pictures in the worksheet.</param>
    /// <exception cref="ArgumentException">Thrown when the image index is out of range.</exception>
    public static void ValidateImageIndex(int imageIndex, int picturesCount)
    {
        if (imageIndex < 0 || imageIndex >= picturesCount)
            throw new ArgumentException(
                $"Image index {imageIndex} is out of range. Worksheet has {picturesCount} images (valid indices: 0-{picturesCount - 1}).");
    }
}
