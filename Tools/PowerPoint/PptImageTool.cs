using System.Drawing;
using System.Drawing.Imaging;
using System.Security.Cryptography;
using System.Text.Json;
using System.Text.Json.Nodes;
using Aspose.Slides;
using Aspose.Slides.Export;
using AsposeMcpServer.Core;

namespace AsposeMcpServer.Tools.PowerPoint;

/// <summary>
///     Unified tool for managing PowerPoint images.
///     Supports: add, edit, delete, get, export_slides, extract
/// </summary>
public class PptImageTool : IAsposeTool
{
    public string Description =>
        @"Manage PowerPoint images. Supports 6 operations: add, edit, delete, get, export_slides, extract.

Usage examples:
- Add image: ppt_image(operation='add', path='presentation.pptx', slideIndex=0, imagePath='image.png', x=100, y=100)
- Edit image: ppt_image(operation='edit', path='presentation.pptx', slideIndex=0, imageIndex=0, width=300, height=200)
- Edit with compression: ppt_image(operation='edit', path='presentation.pptx', slideIndex=0, imageIndex=0, imagePath='new.png', jpegQuality=80, maxWidth=800)
- Delete image: ppt_image(operation='delete', path='presentation.pptx', slideIndex=0, imageIndex=0)
- Get image info: ppt_image(operation='get', path='presentation.pptx', slideIndex=0)
- Export slides as images: ppt_image(operation='export_slides', path='presentation.pptx', outputDir='images/', slideIndexes='0,2,4')
- Extract embedded images: ppt_image(operation='extract', path='presentation.pptx', outputDir='images/', skipDuplicates=true)";

    public object InputSchema => new
    {
        type = "object",
        properties = new
        {
            operation = new
            {
                type = "string",
                description = @"Operation to perform.
- 'add': Add image to slide (required: path, slideIndex, imagePath)
- 'edit': Edit image properties with optional compression/resize (required: path, slideIndex, imageIndex)
- 'delete': Delete image from slide (required: path, slideIndex, imageIndex)
- 'get': Get image information (required: path, slideIndex)
- 'export_slides': Export slides as image files (required: path)
- 'extract': Extract embedded images from presentation (required: path)",
                @enum = new[] { "add", "edit", "delete", "get", "export_slides", "extract" }
            },
            path = new
            {
                type = "string",
                description = "Presentation file path (required for all operations)"
            },
            slideIndex = new
            {
                type = "number",
                description = "Slide index (0-based, required for add/edit/delete/get)"
            },
            imageIndex = new
            {
                type = "number",
                description =
                    "Image index on the slide (0-based, required for edit/delete). Refers to N-th image on slide, not absolute shape index."
            },
            imagePath = new
            {
                type = "string",
                description = "Image file path (required for add, optional for edit)"
            },
            x = new
            {
                type = "number",
                description = "X position in points (optional for add/edit, default: 100)"
            },
            y = new
            {
                type = "number",
                description = "Y position in points (optional for add/edit, default: 100)"
            },
            width = new
            {
                type = "number",
                description = "Width in points (optional for add/edit)"
            },
            height = new
            {
                type = "number",
                description = "Height in points (optional for add/edit)"
            },
            jpegQuality = new
            {
                type = "number",
                description = "JPEG quality 10-100 (optional for edit, re-encode image as JPEG)"
            },
            maxWidth = new
            {
                type = "number",
                description = "Maximum width in pixels for resize (optional for edit)"
            },
            maxHeight = new
            {
                type = "number",
                description = "Maximum height in pixels for resize (optional for edit)"
            },
            outputDir = new
            {
                type = "string",
                description = "Output directory (required for export_slides/extract)"
            },
            format = new
            {
                type = "string",
                description = "Image format: png|jpeg (optional for export_slides/extract, default: png)"
            },
            scale = new
            {
                type = "number",
                description = "Scaling factor (optional for export_slides, default: 1.0)"
            },
            slideIndexes = new
            {
                type = "string",
                description =
                    "Comma-separated slide indexes to export (optional for export_slides, e.g., '0,2,4')"
            },
            skipDuplicates = new
            {
                type = "boolean",
                description = "Skip duplicate images based on content hash (optional for extract, default: false)"
            },
            outputPath = new
            {
                type = "string",
                description = "Output file path (optional for add/edit/delete, defaults to input path)"
            }
        },
        required = new[] { "operation", "path" }
    };

    /// <summary>
    ///     Executes the tool operation with the provided JSON arguments.
    /// </summary>
    /// <param name="arguments">JSON arguments object containing operation parameters.</param>
    /// <returns>Result message as a string.</returns>
    /// <exception cref="ArgumentException">Thrown when operation is unknown.</exception>
    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var operation = ArgumentHelper.GetString(arguments, "operation");
        var path = ArgumentHelper.GetAndValidatePath(arguments);
        var outputPath = ArgumentHelper.GetAndValidateOutputPath(arguments, path);

        return operation.ToLower() switch
        {
            "add" => await AddImageAsync(path, outputPath, arguments),
            "edit" => await EditImageAsync(path, outputPath, arguments),
            "delete" => await DeleteImageAsync(path, outputPath, arguments),
            "get" => await GetImageInfoAsync(path, arguments),
            "export_slides" => await ExportSlidesAsync(path, arguments),
            "extract" => await ExtractImagesAsync(path, arguments),
            _ => throw new ArgumentException($"Unknown operation: {operation}")
        };
    }

    #region Add Operation

    /// <summary>
    ///     Adds an image to a slide.
    /// </summary>
    /// <param name="path">PowerPoint file path.</param>
    /// <param name="outputPath">Output file path.</param>
    /// <param name="arguments">JSON arguments containing slideIndex, imagePath, optional x, y, width, height.</param>
    /// <returns>Success message.</returns>
    /// <exception cref="ArgumentException">Thrown when slideIndex is out of range.</exception>
    /// <exception cref="FileNotFoundException">Thrown when image file is not found.</exception>
    private Task<string> AddImageAsync(string path, string outputPath, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var slideIndex = ArgumentHelper.GetInt(arguments, "slideIndex");
            var imagePath = ArgumentHelper.GetString(arguments, "imagePath");
            var x = ArgumentHelper.GetFloat(arguments, "x", 100);
            var y = ArgumentHelper.GetFloat(arguments, "y", 100);
            var width = ArgumentHelper.GetFloatNullable(arguments, "width");
            var height = ArgumentHelper.GetFloatNullable(arguments, "height");

            if (!File.Exists(imagePath))
                throw new FileNotFoundException($"Image file not found: {imagePath}");

            using var presentation = new Presentation(path);
            var slide = PowerPointHelper.GetSlide(presentation, slideIndex);

            // Use FileStream for memory efficiency
            IPPImage pictureImage;
            int pixelWidth, pixelHeight;

            using (var fileStream = new FileStream(imagePath, FileMode.Open, FileAccess.Read))
            {
                pictureImage = presentation.Images.AddImage(fileStream);
                pixelWidth = pictureImage.Width;
                pixelHeight = pictureImage.Height;
            }

            // Calculate dimensions maintaining aspect ratio
            var (finalWidth, finalHeight) = CalculateDimensions(width, height, pixelWidth, pixelHeight);

            slide.Shapes.AddPictureFrame(ShapeType.Rectangle, x, y, finalWidth, finalHeight, pictureImage);
            presentation.Save(outputPath, SaveFormat.Pptx);

            return $"Image added to slide {slideIndex}. Output: {outputPath}";
        });
    }

    #endregion

    #region Delete Operation

    /// <summary>
    ///     Deletes an image from a slide.
    /// </summary>
    /// <param name="path">PowerPoint file path.</param>
    /// <param name="outputPath">Output file path.</param>
    /// <param name="arguments">JSON arguments containing slideIndex, imageIndex.</param>
    /// <returns>Success message.</returns>
    /// <exception cref="ArgumentException">Thrown when slideIndex or imageIndex is out of range.</exception>
    private Task<string> DeleteImageAsync(string path, string outputPath, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var slideIndex = ArgumentHelper.GetInt(arguments, "slideIndex");
            var imageIndex = ArgumentHelper.GetInt(arguments, "imageIndex");

            using var presentation = new Presentation(path);
            var slide = PowerPointHelper.GetSlide(presentation, slideIndex);
            var pictures = slide.Shapes.OfType<PictureFrame>().ToList();

            if (imageIndex < 0 || imageIndex >= pictures.Count)
                throw new ArgumentException(
                    $"imageIndex {imageIndex} is out of range. Slide {slideIndex} has {pictures.Count} image(s).");

            var pictureFrame = pictures[imageIndex];
            slide.Shapes.Remove(pictureFrame);

            presentation.Save(outputPath, SaveFormat.Pptx);
            return $"Image {imageIndex} deleted from slide {slideIndex}. Output: {outputPath}";
        });
    }

    #endregion

    #region Get Operation

    /// <summary>
    ///     Gets image information from a slide.
    /// </summary>
    /// <param name="path">PowerPoint file path.</param>
    /// <param name="arguments">JSON arguments containing slideIndex.</param>
    /// <returns>JSON string with image information.</returns>
    /// <exception cref="ArgumentException">Thrown when slideIndex is out of range.</exception>
    private Task<string> GetImageInfoAsync(string path, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var slideIndex = ArgumentHelper.GetInt(arguments, "slideIndex");

            using var presentation = new Presentation(path);
            var slide = PowerPointHelper.GetSlide(presentation, slideIndex);
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
                slideIndex,
                imageCount = pictures.Count,
                images = imageInfoList
            };

            return JsonSerializer.Serialize(result, new JsonSerializerOptions { WriteIndented = true });
        });
    }

    #endregion

    #region Export Slides Operation

    /// <summary>
    ///     Exports slides as image files.
    /// </summary>
    /// <param name="path">PowerPoint file path.</param>
    /// <param name="arguments">JSON arguments containing outputDir, optional slideIndexes, format, scale.</param>
    /// <returns>Success message with exported count.</returns>
    /// <exception cref="ArgumentException">Thrown when slideIndexes contains invalid index.</exception>
    private Task<string> ExportSlidesAsync(string path, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var outputDir = ArgumentHelper.GetStringNullable(arguments, "outputDir") ??
                            Path.GetDirectoryName(path) ?? ".";
            var formatStr = ArgumentHelper.GetString(arguments, "format", "png");
            var scale = ArgumentHelper.GetFloat(arguments, "scale", 1.0f);
            var slideIndexesStr = ArgumentHelper.GetStringNullable(arguments, "slideIndexes");

#pragma warning disable CA1416 // Validate platform compatibility
            var format = formatStr.ToLower() switch
            {
                "jpeg" or "jpg" => ImageFormat.Jpeg,
                _ => ImageFormat.Png
            };
            var extension = format.Equals(ImageFormat.Png) ? "png" : "jpg";
#pragma warning restore CA1416 // Validate platform compatibility

            Directory.CreateDirectory(outputDir);

            using var presentation = new Presentation(path);
            var slideIndexes = ParseSlideIndexes(slideIndexesStr, presentation.Slides.Count);

            var exportedCount = 0;
            foreach (var i in slideIndexes)
            {
                using var bmp = presentation.Slides[i].GetThumbnail(scale, scale);
                var fileName = Path.Combine(outputDir, $"slide_{i + 1}.{extension}");
#pragma warning disable CA1416 // Validate platform compatibility
                bmp.Save(fileName, format);
#pragma warning restore CA1416 // Validate platform compatibility
                exportedCount++;
            }

            return $"Exported {exportedCount} slides. Output: {Path.GetFullPath(outputDir)}";
        });
    }

    #endregion

    #region Extract Operation

    /// <summary>
    ///     Extracts embedded images from the presentation.
    /// </summary>
    /// <param name="path">PowerPoint file path.</param>
    /// <param name="arguments">JSON arguments containing outputDir, optional format, skipDuplicates.</param>
    /// <returns>Success message with extracted count.</returns>
    private Task<string> ExtractImagesAsync(string path, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var outputDir = ArgumentHelper.GetStringNullable(arguments, "outputDir") ??
                            Path.GetDirectoryName(path) ?? ".";
            var formatStr = ArgumentHelper.GetString(arguments, "format", "png");
            var skipDuplicates = ArgumentHelper.GetBool(arguments, "skipDuplicates", false);

#pragma warning disable CA1416 // Validate platform compatibility
            var format = formatStr.ToLower() switch
            {
                "jpeg" or "jpg" => ImageFormat.Jpeg,
                _ => ImageFormat.Png
            };
            var extension = format.Equals(ImageFormat.Png) ? "png" : "jpg";
#pragma warning restore CA1416 // Validate platform compatibility

            Directory.CreateDirectory(outputDir);

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

                        var fileName = Path.Combine(outputDir, $"slide{slideNum}_img{++count}.{extension}");
                        var systemImage = image.SystemImage;
#pragma warning disable CA1416 // Validate platform compatibility
                        systemImage.Save(fileName, format);
#pragma warning restore CA1416 // Validate platform compatibility
                    }
            }

            var result = $"Extracted {count} images. Output: {Path.GetFullPath(outputDir)}";
            if (skipDuplicates && skippedCount > 0)
                result += $" (skipped {skippedCount} duplicates)";

            return result;
        });
    }

    #endregion

    #region Edit Operation

    /// <summary>
    ///     Edits image properties with optional compression and resize.
    /// </summary>
    /// <param name="path">PowerPoint file path.</param>
    /// <param name="outputPath">Output file path.</param>
    /// <param name="arguments">
    ///     JSON arguments containing slideIndex, imageIndex, optional x, y, width, height, imagePath,
    ///     jpegQuality, maxWidth, maxHeight.
    /// </param>
    /// <returns>Success message.</returns>
    /// <exception cref="ArgumentException">Thrown when slideIndex or imageIndex is out of range.</exception>
    private Task<string> EditImageAsync(string path, string outputPath, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var slideIndex = ArgumentHelper.GetInt(arguments, "slideIndex");
            var imageIndex = ArgumentHelper.GetInt(arguments, "imageIndex");
            var imagePath = ArgumentHelper.GetStringNullable(arguments, "imagePath");
            var x = ArgumentHelper.GetFloatNullable(arguments, "x");
            var y = ArgumentHelper.GetFloatNullable(arguments, "y");
            var width = ArgumentHelper.GetFloatNullable(arguments, "width");
            var height = ArgumentHelper.GetFloatNullable(arguments, "height");
            var jpegQuality = ArgumentHelper.GetIntNullable(arguments, "jpegQuality");
            var maxWidth = ArgumentHelper.GetIntNullable(arguments, "maxWidth");
            var maxHeight = ArgumentHelper.GetIntNullable(arguments, "maxHeight");

            using var presentation = new Presentation(path);
            var slide = PowerPointHelper.GetSlide(presentation, slideIndex);
            var pictures = slide.Shapes.OfType<PictureFrame>().ToList();

            if (imageIndex < 0 || imageIndex >= pictures.Count)
                throw new ArgumentException(
                    $"imageIndex {imageIndex} is out of range. Slide {slideIndex} has {pictures.Count} image(s).");

            var pictureFrame = pictures[imageIndex];
            var changes = new List<string>();

            // Replace image if provided
            if (!string.IsNullOrEmpty(imagePath) && File.Exists(imagePath))
            {
                var newImage = ProcessAndAddImage(presentation, imagePath, jpegQuality, maxWidth, maxHeight,
                    out var processingDetails);
                pictureFrame.PictureFormat.Picture.Image = newImage;
                changes.AddRange(processingDetails);
            }

            // Update position
            if (x.HasValue)
            {
                pictureFrame.X = x.Value;
                changes.Add($"x={x.Value}");
            }

            if (y.HasValue)
            {
                pictureFrame.Y = y.Value;
                changes.Add($"y={y.Value}");
            }

            // Update size
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

            presentation.Save(outputPath, SaveFormat.Pptx);

            var changesStr = changes.Count > 0 ? string.Join(", ", changes) : "no changes";
            return $"Image {imageIndex} on slide {slideIndex} updated ({changesStr}). Output: {outputPath}";
        });
    }

    /// <summary>
    ///     Processes image with optional compression/resize and adds to presentation.
    /// </summary>
    private IPPImage ProcessAndAddImage(IPresentation presentation, string imagePath, int? jpegQuality, int? maxWidth,
        int? maxHeight, out List<string> processingDetails)
    {
        processingDetails = new List<string>();

        if (jpegQuality.HasValue || maxWidth.HasValue || maxHeight.HasValue)
        {
#pragma warning disable CA1416 // Validate platform compatibility
            using var fileStream = new FileStream(imagePath, FileMode.Open, FileAccess.Read);
            using var src = Image.FromStream(fileStream);

            var processedImage = src;
            var needsDispose = false;

            // Resize if needed
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

            // Encode to JPEG with quality or PNG
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

        // No processing needed, use FileStream directly
        using var fs = new FileStream(imagePath, FileMode.Open, FileAccess.Read);
        processingDetails.Add("image replaced");
        return presentation.Images.AddImage(fs);
    }

    #endregion

    #region Helper Methods

    /// <summary>
    ///     Calculates final dimensions maintaining aspect ratio.
    /// </summary>
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

        // Default: 300 points width, maintain aspect ratio
        var defaultWidth = 300f;
        var defaultRatio = pixelWidth > 0 ? (float)pixelHeight / pixelWidth : 1;
        return (defaultWidth, defaultWidth * defaultRatio);
    }

    /// <summary>
    ///     Calculates new image size maintaining aspect ratio within max bounds.
    /// </summary>
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
    private static List<int> ParseSlideIndexes(string? slideIndexesStr, int totalSlides)
    {
        if (string.IsNullOrWhiteSpace(slideIndexesStr))
            return Enumerable.Range(0, totalSlides).ToList();

        var indexes = new List<int>();
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
    private static string ComputeImageHash(byte[] data)
    {
        var hashBytes = MD5.HashData(data);
        return Convert.ToHexString(hashBytes);
    }

    #endregion
}