using System.Text.Json.Nodes;
using Aspose.Slides;
using Aspose.Slides.Export;
using AsposeMcpServer.Core;

namespace AsposeMcpServer.Tools.PowerPoint;

/// <summary>
///     Unified tool for managing PowerPoint images (add, edit)
///     Merges: PptAddImageTool, PptEditImageTool
/// </summary>
public class PptImageTool : IAsposeTool
{
    public string Description => @"Manage PowerPoint images. Supports 2 operations: add, edit.

Usage examples:
- Add image: ppt_image(operation='add', path='presentation.pptx', slideIndex=0, imagePath='image.png', x=100, y=100, width=200, height=150)
- Edit image: ppt_image(operation='edit', path='presentation.pptx', slideIndex=0, shapeIndex=0, width=300, height=200)

Note: shapeIndex for edit operation refers to the image index (0-based) among all images on the slide, not the absolute shape index.";

    public object InputSchema => new
    {
        type = "object",
        properties = new
        {
            operation = new
            {
                type = "string",
                description = @"Operation to perform.
- 'add': Add an image (required params: path, slideIndex, imagePath)
- 'edit': Edit an image (required params: path, slideIndex, shapeIndex)",
                @enum = new[] { "add", "edit" }
            },
            path = new
            {
                type = "string",
                description = "Presentation file path (required for all operations)"
            },
            slideIndex = new
            {
                type = "number",
                description = "Slide index (0-based)"
            },
            shapeIndex = new
            {
                type = "number",
                description =
                    "Image index (0-based, required for edit). This refers to the N-th image on the slide, not the absolute shape index. Use 0 for the first image, 1 for the second image, etc."
            },
            imagePath = new
            {
                type = "string",
                description = "Image file path (required for add, optional for edit)"
            },
            x = new
            {
                type = "number",
                description = "X position (optional, default: 100 for add)"
            },
            y = new
            {
                type = "number",
                description = "Y position (optional, default: 100 for add)"
            },
            width = new
            {
                type = "number",
                description = "Image width (optional)"
            },
            height = new
            {
                type = "number",
                description = "Image height (optional)"
            },
            outputPath = new
            {
                type = "string",
                description = "Output file path (optional, for add/edit/delete operations, defaults to input path)"
            }
        },
        required = new[] { "operation", "path", "slideIndex" }
    };

    /// <summary>
    ///     Executes the tool operation with the provided JSON arguments
    /// </summary>
    /// <param name="arguments">JSON arguments object containing operation parameters</param>
    /// <returns>Result message as a string</returns>
    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var operation = ArgumentHelper.GetString(arguments, "operation");
        var path = ArgumentHelper.GetAndValidatePath(arguments);
        var outputPath = ArgumentHelper.GetAndValidateOutputPath(arguments, path);
        SecurityHelper.ValidateFilePath(path, allowAbsolutePaths: true);
        var slideIndex = ArgumentHelper.GetInt(arguments, "slideIndex");

        return operation.ToLower() switch
        {
            "add" => await AddImageAsync(path, outputPath, slideIndex, arguments),
            "edit" => await EditImageAsync(path, outputPath, slideIndex, arguments),
            _ => throw new ArgumentException($"Unknown operation: {operation}")
        };
    }

    /// <summary>
    ///     Adds an image to a slide
    /// </summary>
    /// <param name="path">PowerPoint file path</param>
    /// <param name="outputPath">Output file path</param>
    /// <param name="slideIndex">Slide index (0-based)</param>
    /// <param name="arguments">JSON arguments containing imagePath, optional x, y, width, height</param>
    /// <returns>Success message</returns>
    private Task<string> AddImageAsync(string path, string outputPath, int slideIndex, JsonObject? arguments)
    {
        return Task.Run(async () =>
        {
            var imagePath = ArgumentHelper.GetString(arguments, "imagePath");
            var x = ArgumentHelper.GetFloat(arguments, "x", "x", false, 100);
            var y = ArgumentHelper.GetFloat(arguments, "y", "y", false, 100);
            var width = ArgumentHelper.GetFloatNullable(arguments, "width");
            var height = ArgumentHelper.GetFloatNullable(arguments, "height");

            if (!File.Exists(imagePath)) throw new FileNotFoundException($"Image file not found: {imagePath}");

            using var presentation = new Presentation(path);
            var slide = PowerPointHelper.GetSlide(presentation, slideIndex);

            // Read image bytes asynchronously
            var imageBytes = await File.ReadAllBytesAsync(imagePath);

            // Check if image already exists in presentation to avoid duplication
            IPPImage? existingImage = null;
            foreach (var img in presentation.Images)
                // Compare image sizes as a simple check (in production, consider using hash comparison)
                if (img.BinaryData.Length == imageBytes.Length)
                    // Additional check: compare first and last bytes for quick verification
                    if (img.BinaryData.Length > 0 &&
                        img.BinaryData[0] == imageBytes[0] &&
                        img.BinaryData[^1] == imageBytes[^1])
                    {
                        existingImage = img;
                        break;
                    }

            // Use existing image if found, otherwise add new one
            var pictureImage = existingImage ?? presentation.Images.AddImage(imageBytes);

            // Handle width and height with proper unit conversion
            // Note: pictureImage.Width/Height are in pixels, but PowerPoint uses Points (1/72 inch)
            // Default width: 300 points (approximately 4.17 inches)
            float finalWidth;
            float finalHeight;

            if (width.HasValue && height.HasValue)
            {
                finalWidth = width.Value;
                finalHeight = height.Value;
            }
            else if (width.HasValue)
            {
                finalWidth = width.Value;
                // Calculate height maintaining aspect ratio
                // Convert pixel dimensions to points (assuming 96 DPI: 1 pixel = 0.75 points)
                var pixelWidth = pictureImage.Width;
                var pixelHeight = pictureImage.Height;
                if (pixelWidth > 0)
                    finalHeight = finalWidth * ((float)pixelHeight / pixelWidth);
                else
                    finalHeight = finalWidth; // Fallback to square if width is 0
            }
            else if (height.HasValue)
            {
                finalHeight = height.Value;
                // Calculate width maintaining aspect ratio
                var pixelWidth = pictureImage.Width;
                var pixelHeight = pictureImage.Height;
                if (pixelHeight > 0)
                    finalWidth = finalHeight * ((float)pixelWidth / pixelHeight);
                else
                    finalWidth = finalHeight; // Fallback to square if height is 0
            }
            else
            {
                // Default size: 300 points width, maintain aspect ratio
                finalWidth = 300;
                var pixelWidth = pictureImage.Width;
                var pixelHeight = pictureImage.Height;
                if (pixelWidth > 0)
                    finalHeight = finalWidth * ((float)pixelHeight / pixelWidth);
                else
                    finalHeight = 300; // Fallback to square
            }

            slide.Shapes.AddPictureFrame(ShapeType.Rectangle, x, y, finalWidth, finalHeight, pictureImage);

            presentation.Save(outputPath, SaveFormat.Pptx);

            return $"Image added to slide {slideIndex}. Output: {outputPath}";
        });
    }

    /// <summary>
    ///     Edits image properties
    /// </summary>
    /// <param name="path">PowerPoint file path</param>
    /// <param name="outputPath">Output file path</param>
    /// <param name="slideIndex">Slide index (0-based)</param>
    /// <param name="arguments">JSON arguments containing shapeIndex (image index), optional x, y, width, height, imagePath</param>
    /// <returns>Success message</returns>
    private Task<string> EditImageAsync(string path, string outputPath, int slideIndex, JsonObject? arguments)
    {
        return Task.Run(async () =>
        {
            var imageIndex = ArgumentHelper.GetInt(arguments, "shapeIndex");
            var imagePath = ArgumentHelper.GetStringNullable(arguments, "imagePath");
            var x = ArgumentHelper.GetFloatNullable(arguments, "x");
            var y = ArgumentHelper.GetFloatNullable(arguments, "y");
            var width = ArgumentHelper.GetFloatNullable(arguments, "width");
            var height = ArgumentHelper.GetFloatNullable(arguments, "height");

            using var presentation = new Presentation(path);
            var slide = PowerPointHelper.GetSlide(presentation, slideIndex);

            var pictures = slide.Shapes.OfType<PictureFrame>().ToList();

            if (imageIndex < 0 || imageIndex >= pictures.Count)
            {
                var totalShapes = slide.Shapes.Count;
                var totalImages = pictures.Count;
                throw new ArgumentException(
                    $"Image index {imageIndex} is out of range. Slide {slideIndex} has {totalImages} image(s) " +
                    $"(out of {totalShapes} total shape(s)). Valid image indices: 0 to {totalImages - 1}. " +
                    $"Note: imageIndex refers to the N-th image on the slide, not the absolute shape index.");
            }

            var pictureFrame = pictures[imageIndex];

            if (!string.IsNullOrEmpty(imagePath) && File.Exists(imagePath))
            {
                var imageBytes = await File.ReadAllBytesAsync(imagePath);

                IPPImage? existingImage = null;
                foreach (var img in presentation.Images)
                    if (img.BinaryData.Length == imageBytes.Length)
                        if (img.BinaryData.Length > 0 &&
                            img.BinaryData[0] == imageBytes[0] &&
                            img.BinaryData[^1] == imageBytes[^1])
                        {
                            existingImage = img;
                            break;
                        }

                var newImage = existingImage ?? presentation.Images.AddImage(imageBytes);
                pictureFrame.PictureFormat.Picture.Image = newImage;
            }

            if (x.HasValue) pictureFrame.X = x.Value;
            if (y.HasValue) pictureFrame.Y = y.Value;
            if (width.HasValue) pictureFrame.Width = width.Value;
            if (height.HasValue) pictureFrame.Height = height.Value;

            presentation.Save(outputPath, SaveFormat.Pptx);
            return $"Image {imageIndex} on slide {slideIndex} updated. Output: {outputPath}";
        });
    }
}