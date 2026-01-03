using System.ComponentModel;
using System.Drawing;
using System.Drawing.Imaging;
using System.Security.Cryptography;
using System.Text.Json;
using Aspose.Slides;
using AsposeMcpServer.Core.Helpers;
using AsposeMcpServer.Core.Session;
using ModelContextProtocol.Server;

namespace AsposeMcpServer.Tools.PowerPoint;

/// <summary>
///     Unified tool for managing PowerPoint images.
///     Supports: add, edit, delete, get, export_slides, extract
/// </summary>
[McpServerToolType]
public class PptImageTool
{
    /// <summary>
    ///     Session manager for document lifecycle management.
    /// </summary>
    private readonly DocumentSessionManager? _sessionManager;

    /// <summary>
    ///     Initializes a new instance of the <see cref="PptImageTool" /> class.
    /// </summary>
    /// <param name="sessionManager">Optional session manager for in-memory document editing.</param>
    public PptImageTool(DocumentSessionManager? sessionManager = null)
    {
        _sessionManager = sessionManager;
    }

    [McpServerTool(Name = "ppt_image")]
    [Description(@"Manage PowerPoint images. Supports 6 operations: add, edit, delete, get, export_slides, extract.

Usage examples:
- Add image: ppt_image(operation='add', path='presentation.pptx', slideIndex=0, imagePath='image.png', x=100, y=100)
- Edit image: ppt_image(operation='edit', path='presentation.pptx', slideIndex=0, imageIndex=0, width=300, height=200)
- Edit with compression: ppt_image(operation='edit', path='presentation.pptx', slideIndex=0, imageIndex=0, imagePath='new.png', jpegQuality=80, maxWidth=800)
- Delete image: ppt_image(operation='delete', path='presentation.pptx', slideIndex=0, imageIndex=0)
- Get image info: ppt_image(operation='get', path='presentation.pptx', slideIndex=0)
- Export slides as images: ppt_image(operation='export_slides', path='presentation.pptx', outputDir='images/', slideIndexes='0,2,4')
- Extract embedded images: ppt_image(operation='extract', path='presentation.pptx', outputDir='images/', skipDuplicates=true)")]
    public string Execute(
        [Description("Operation: add, edit, delete, get, export_slides, extract")]
        string operation,
        [Description("Presentation file path (required if no sessionId)")]
        string? path = null,
        [Description("Session ID for in-memory editing")]
        string? sessionId = null,
        [Description("Output file path (file mode only)")]
        string? outputPath = null,
        [Description("Slide index (0-based, required for add/edit/delete/get)")]
        int? slideIndex = null,
        [Description("Image index on the slide (0-based, required for edit/delete)")]
        int? imageIndex = null,
        [Description("Image file path (required for add, optional for edit)")]
        string? imagePath = null,
        [Description("X position in points (optional for add/edit, default: 100)")]
        float x = 100,
        [Description("Y position in points (optional for add/edit, default: 100)")]
        float y = 100,
        [Description("Width in points (optional for add/edit)")]
        float? width = null,
        [Description("Height in points (optional for add/edit)")]
        float? height = null,
        [Description("JPEG quality 10-100 (optional for edit, re-encode image as JPEG)")]
        int? jpegQuality = null,
        [Description("Maximum width in pixels for resize (optional for edit)")]
        int? maxWidth = null,
        [Description("Maximum height in pixels for resize (optional for edit)")]
        int? maxHeight = null,
        [Description("Output directory (required for export_slides/extract)")]
        string? outputDir = null,
        [Description("Image format: png|jpeg (optional for export_slides/extract, default: png)")]
        string format = "png",
        [Description("Scaling factor (optional for export_slides, default: 1.0)")]
        float scale = 1.0f,
        [Description("Comma-separated slide indexes to export (optional for export_slides, e.g., '0,2,4')")]
        string? slideIndexes = null,
        [Description("Skip duplicate images based on content hash (optional for extract, default: false)")]
        bool skipDuplicates = false)
    {
        if (operation.ToLower() == "export_slides")
            return ExportSlides(path!, outputDir, slideIndexes, format, scale);
        if (operation.ToLower() == "extract")
            return ExtractImages(path!, outputDir, format, skipDuplicates);

        using var ctx = DocumentContext<Presentation>.Create(_sessionManager, sessionId, path);

        return operation.ToLower() switch
        {
            "add" => AddImage(ctx, outputPath, slideIndex, imagePath, x, y, width, height),
            "edit" => EditImage(ctx, outputPath, slideIndex, imageIndex, imagePath, x, y, width, height, jpegQuality,
                maxWidth, maxHeight),
            "delete" => DeleteImage(ctx, outputPath, slideIndex, imageIndex),
            "get" => GetImageInfo(ctx, slideIndex),
            _ => throw new ArgumentException($"Unknown operation: {operation}")
        };
    }

    #region Add Operation

    /// <summary>
    ///     Adds an image to a slide.
    /// </summary>
    /// <param name="ctx">The document context.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="slideIndex">The zero-based index of the slide.</param>
    /// <param name="imagePath">The path to the image file.</param>
    /// <param name="x">The X position in points.</param>
    /// <param name="y">The Y position in points.</param>
    /// <param name="width">The width in points (optional).</param>
    /// <param name="height">The height in points (optional).</param>
    /// <returns>A message indicating the result of the operation.</returns>
    /// <exception cref="ArgumentException">Thrown when slideIndex or imagePath is not provided.</exception>
    /// <exception cref="FileNotFoundException">Thrown when the image file does not exist.</exception>
    private static string AddImage(DocumentContext<Presentation> ctx, string? outputPath,
        int? slideIndex, string? imagePath, float x, float y, float? width, float? height)
    {
        if (!slideIndex.HasValue)
            throw new ArgumentException("slideIndex is required for add operation");
        if (string.IsNullOrEmpty(imagePath))
            throw new ArgumentException("imagePath is required for add operation");
        if (!File.Exists(imagePath))
            throw new FileNotFoundException($"Image file not found: {imagePath}");

        var presentation = ctx.Document;
        var slide = PowerPointHelper.GetSlide(presentation, slideIndex.Value);

        IPPImage pictureImage;
        int pixelWidth, pixelHeight;

        using (var fileStream = new FileStream(imagePath, FileMode.Open, FileAccess.Read))
        {
            pictureImage = presentation.Images.AddImage(fileStream);
            pixelWidth = pictureImage.Width;
            pixelHeight = pictureImage.Height;
        }

        var (finalWidth, finalHeight) = CalculateDimensions(width, height, pixelWidth, pixelHeight);

        slide.Shapes.AddPictureFrame(ShapeType.Rectangle, x, y, finalWidth, finalHeight, pictureImage);

        ctx.Save(outputPath);

        var result = $"Image added to slide {slideIndex}. ";
        result += ctx.GetOutputMessage(outputPath);
        return result;
    }

    #endregion

    #region Delete Operation

    /// <summary>
    ///     Deletes an image from a slide.
    /// </summary>
    /// <param name="ctx">The document context.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="slideIndex">The zero-based index of the slide.</param>
    /// <param name="imageIndex">The zero-based index of the image on the slide.</param>
    /// <returns>A message indicating the result of the operation.</returns>
    /// <exception cref="ArgumentException">Thrown when slideIndex or imageIndex is not provided or out of range.</exception>
    private static string DeleteImage(DocumentContext<Presentation> ctx, string? outputPath,
        int? slideIndex, int? imageIndex)
    {
        if (!slideIndex.HasValue)
            throw new ArgumentException("slideIndex is required for delete operation");
        if (!imageIndex.HasValue)
            throw new ArgumentException("imageIndex is required for delete operation");

        var presentation = ctx.Document;
        var slide = PowerPointHelper.GetSlide(presentation, slideIndex.Value);
        var pictures = slide.Shapes.OfType<PictureFrame>().ToList();

        if (imageIndex.Value < 0 || imageIndex.Value >= pictures.Count)
            throw new ArgumentException(
                $"imageIndex {imageIndex.Value} is out of range. Slide {slideIndex.Value} has {pictures.Count} image(s).");

        var pictureFrame = pictures[imageIndex.Value];
        slide.Shapes.Remove(pictureFrame);

        ctx.Save(outputPath);

        var result = $"Image {imageIndex} deleted from slide {slideIndex}. ";
        result += ctx.GetOutputMessage(outputPath);
        return result;
    }

    #endregion

    #region Get Operation

    /// <summary>
    ///     Gets image information from a slide.
    /// </summary>
    /// <param name="ctx">The document context.</param>
    /// <param name="slideIndex">The zero-based index of the slide.</param>
    /// <returns>A JSON string containing image information including count, positions, and sizes.</returns>
    /// <exception cref="ArgumentException">Thrown when slideIndex is not provided.</exception>
    private static string GetImageInfo(DocumentContext<Presentation> ctx, int? slideIndex)
    {
        if (!slideIndex.HasValue)
            throw new ArgumentException("slideIndex is required for get operation");

        var presentation = ctx.Document;
        var slide = PowerPointHelper.GetSlide(presentation, slideIndex.Value);
        var pictures = slide.Shapes.OfType<PictureFrame>().ToList();

        var imageInfoList = pictures.Select((pic, index) => new
        {
            imageIndex = index,
            x = pic.X,
            y = pic.Y,
            width = pic.Width,
            height = pic.Height,
            contentType = pic.PictureFormat.Picture.Image?.ContentType ?? "unknown"
        }).ToList();

        var result = new
        {
            slideIndex = slideIndex.Value,
            imageCount = pictures.Count,
            images = imageInfoList
        };

        return JsonSerializer.Serialize(result, new JsonSerializerOptions { WriteIndented = true });
    }

    #endregion

    #region Export Slides Operation

    /// <summary>
    ///     Exports slides as image files.
    /// </summary>
    /// <param name="path">The presentation file path.</param>
    /// <param name="outputDir">The output directory path.</param>
    /// <param name="slideIndexesStr">Comma-separated slide indexes to export (optional).</param>
    /// <param name="formatStr">The image format (png or jpeg).</param>
    /// <param name="scale">The scaling factor.</param>
    /// <returns>A message indicating the result of the operation.</returns>
    /// <exception cref="ArgumentException">Thrown when slide indexes are invalid.</exception>
    private static string ExportSlides(string path, string? outputDir, string? slideIndexesStr, string formatStr,
        float scale)
    {
        SecurityHelper.ValidateFilePath(path, "path", true);

        var actualOutputDir = outputDir ?? Path.GetDirectoryName(path) ?? ".";

#pragma warning disable CA1416 // Validate platform compatibility
        var format = formatStr.ToLower() switch
        {
            "jpeg" or "jpg" => ImageFormat.Jpeg,
            _ => ImageFormat.Png
        };
        var extension = format.Equals(ImageFormat.Png) ? "png" : "jpg";
#pragma warning restore CA1416 // Validate platform compatibility

        Directory.CreateDirectory(actualOutputDir);

        using var presentation = new Presentation(path);
        var slideIndexList = ParseSlideIndexes(slideIndexesStr, presentation.Slides.Count);

        var exportedCount = 0;
        foreach (var i in slideIndexList)
        {
            using var bmp = presentation.Slides[i].GetThumbnail(scale, scale);
            var fileName = Path.Combine(actualOutputDir, $"slide_{i + 1}.{extension}");
#pragma warning disable CA1416 // Validate platform compatibility
            bmp.Save(fileName, format);
#pragma warning restore CA1416 // Validate platform compatibility
            exportedCount++;
        }

        return $"Exported {exportedCount} slides. Output: {Path.GetFullPath(actualOutputDir)}";
    }

    #endregion

    #region Extract Operation

    /// <summary>
    ///     Extracts embedded images from the presentation.
    /// </summary>
    /// <param name="path">The presentation file path.</param>
    /// <param name="outputDir">The output directory path.</param>
    /// <param name="formatStr">The image format (png or jpeg).</param>
    /// <param name="skipDuplicates">True to skip duplicate images based on content hash.</param>
    /// <returns>A message indicating the result of the operation.</returns>
    private static string ExtractImages(string path, string? outputDir, string formatStr, bool skipDuplicates)
    {
        SecurityHelper.ValidateFilePath(path, "path", true);

        var actualOutputDir = outputDir ?? Path.GetDirectoryName(path) ?? ".";

#pragma warning disable CA1416 // Validate platform compatibility
        var format = formatStr.ToLower() switch
        {
            "jpeg" or "jpg" => ImageFormat.Jpeg,
            _ => ImageFormat.Png
        };
        var extension = format.Equals(ImageFormat.Png) ? "png" : "jpg";
#pragma warning restore CA1416 // Validate platform compatibility

        Directory.CreateDirectory(actualOutputDir);

        using var presentation = new Presentation(path);
        var count = 0;
        var skippedCount = 0;
        var exportedHashes = new HashSet<string>();

        var slideNum = 0;
        foreach (var slide in presentation.Slides)
        {
            slideNum++;
            foreach (var shape in slide.Shapes)
                if (shape is PictureFrame { PictureFormat.Picture.Image: not null } pic)
                {
                    var image = pic.PictureFormat.Picture.Image;

                    if (skipDuplicates)
                    {
                        var hash = ComputeImageHash(image.BinaryData);
                        if (!exportedHashes.Add(hash))
                        {
                            skippedCount++;
                            continue;
                        }
                    }

                    var fileName = Path.Combine(actualOutputDir, $"slide{slideNum}_img{++count}.{extension}");
                    var systemImage = image.SystemImage;
#pragma warning disable CA1416 // Validate platform compatibility
                    systemImage.Save(fileName, format);
#pragma warning restore CA1416 // Validate platform compatibility
                }
        }

        var result = $"Extracted {count} images. Output: {Path.GetFullPath(actualOutputDir)}";
        if (skipDuplicates && skippedCount > 0)
            result += $" (skipped {skippedCount} duplicates)";

        return result;
    }

    #endregion

    #region Edit Operation

    /// <summary>
    ///     Edits image properties with optional compression and resize.
    /// </summary>
    /// <param name="ctx">The document context.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="slideIndex">The zero-based index of the slide.</param>
    /// <param name="imageIndex">The zero-based index of the image on the slide.</param>
    /// <param name="imagePath">The path to a new image file (optional).</param>
    /// <param name="x">The X position in points.</param>
    /// <param name="y">The Y position in points.</param>
    /// <param name="width">The width in points (optional).</param>
    /// <param name="height">The height in points (optional).</param>
    /// <param name="jpegQuality">The JPEG quality 10-100 (optional).</param>
    /// <param name="maxWidth">The maximum width in pixels for resize (optional).</param>
    /// <param name="maxHeight">The maximum height in pixels for resize (optional).</param>
    /// <returns>A message indicating the result of the operation.</returns>
    /// <exception cref="ArgumentException">Thrown when slideIndex or imageIndex is not provided or out of range.</exception>
    private static string EditImage(DocumentContext<Presentation> ctx, string? outputPath,
        int? slideIndex, int? imageIndex, string? imagePath,
        float x, float y, float? width, float? height,
        int? jpegQuality, int? maxWidth, int? maxHeight)
    {
        if (!slideIndex.HasValue)
            throw new ArgumentException("slideIndex is required for edit operation");
        if (!imageIndex.HasValue)
            throw new ArgumentException("imageIndex is required for edit operation");

        var presentation = ctx.Document;
        var slide = PowerPointHelper.GetSlide(presentation, slideIndex.Value);
        var pictures = slide.Shapes.OfType<PictureFrame>().ToList();

        if (imageIndex.Value < 0 || imageIndex.Value >= pictures.Count)
            throw new ArgumentException(
                $"imageIndex {imageIndex.Value} is out of range. Slide {slideIndex.Value} has {pictures.Count} image(s).");

        var pictureFrame = pictures[imageIndex.Value];
        List<string> changes = [];

        if (!string.IsNullOrEmpty(imagePath) && File.Exists(imagePath))
        {
            var newImage = ProcessAndAddImage(presentation, imagePath, jpegQuality, maxWidth, maxHeight,
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

        ctx.Save(outputPath);

        var changesStr = changes.Count > 0 ? string.Join(", ", changes) : "position updated";
        var result = $"Image {imageIndex} on slide {slideIndex} updated ({changesStr}). ";
        result += ctx.GetOutputMessage(outputPath);
        return result;
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
    private static IPPImage ProcessAndAddImage(IPresentation presentation, string imagePath, int? jpegQuality,
        int? maxWidth,
        int? maxHeight, out List<string> processingDetails)
    {
        processingDetails = [];

        if (jpegQuality.HasValue || maxWidth.HasValue || maxHeight.HasValue)
        {
#pragma warning disable CA1416 // Validate platform compatibility
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
#pragma warning restore CA1416 // Validate platform compatibility

            ms.Position = 0;
            processingDetails.Insert(0, "image replaced");
            return presentation.Images.AddImage(ms);
        }

        using var fs = new FileStream(imagePath, FileMode.Open, FileAccess.Read);
        processingDetails.Add("image replaced");
        return presentation.Images.AddImage(fs);
    }

    #endregion

    #region Helper Methods

    /// <summary>
    ///     Calculates final dimensions maintaining aspect ratio.
    /// </summary>
    /// <param name="width">The requested width (optional).</param>
    /// <param name="height">The requested height (optional).</param>
    /// <param name="pixelWidth">The original image width in pixels.</param>
    /// <param name="pixelHeight">The original image height in pixels.</param>
    /// <returns>A tuple containing the calculated width and height.</returns>
    private static (float width, float height) CalculateDimensions(float? width, float? height, int pixelWidth,
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
    private static Size CalculateResizeSize(int width, int height, int? maxWidth, int? maxHeight)
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
    ///     Parses comma-separated slide indexes string.
    /// </summary>
    /// <param name="slideIndexesStr">Comma-separated slide indexes string.</param>
    /// <param name="totalSlides">The total number of slides in the presentation.</param>
    /// <returns>A list of valid slide indexes.</returns>
    /// <exception cref="ArgumentException">Thrown when a slide index is invalid or out of range.</exception>
    private static List<int> ParseSlideIndexes(string? slideIndexesStr, int totalSlides)
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
    private static string ComputeImageHash(byte[] data)
    {
        var hashBytes = MD5.HashData(data);
        return Convert.ToHexString(hashBytes);
    }

    #endregion
}