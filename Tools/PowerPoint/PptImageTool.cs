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
- Edit image: ppt_image(operation='edit', path='presentation.pptx', slideIndex=0, shapeIndex=0, width=300, height=200)";

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
                description = "Shape index of the image (0-based, required for edit)"
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

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var operation = ArgumentHelper.GetString(arguments, "operation");
        var path = ArgumentHelper.GetAndValidatePath(arguments);
        SecurityHelper.ValidateFilePath(path);
        var slideIndex = ArgumentHelper.GetInt(arguments, "slideIndex");

        return operation.ToLower() switch
        {
            "add" => await AddImageAsync(arguments, path, slideIndex),
            "edit" => await EditImageAsync(arguments, path, slideIndex),
            _ => throw new ArgumentException($"Unknown operation: {operation}")
        };
    }

    /// <summary>
    ///     Adds an image to a slide
    /// </summary>
    /// <param name="arguments">JSON arguments containing imagePath, optional x, y, width, height, outputPath</param>
    /// <param name="path">PowerPoint file path</param>
    /// <param name="slideIndex">Slide index (0-based)</param>
    /// <returns>Success message</returns>
    private async Task<string> AddImageAsync(JsonObject? arguments, string path, int slideIndex)
    {
        var imagePath = ArgumentHelper.GetString(arguments, "imagePath");
        var x = ArgumentHelper.GetFloat(arguments, "x", "x", false, 100);
        var y = ArgumentHelper.GetFloat(arguments, "y", "y", false, 100);
        var width = ArgumentHelper.GetFloatNullable(arguments, "width");
        var height = ArgumentHelper.GetFloatNullable(arguments, "height");

        if (!File.Exists(imagePath)) throw new FileNotFoundException($"Image file not found: {imagePath}");

        using var presentation = new Presentation(path);
        if (slideIndex < 0 || slideIndex >= presentation.Slides.Count)
            throw new ArgumentException($"slideIndex must be between 0 and {presentation.Slides.Count - 1}");

        var slide = presentation.Slides[slideIndex];

        await using var image = File.OpenRead(imagePath);
        var pictureImage = presentation.Images.AddImage(image);

        if (width.HasValue && height.HasValue)
            slide.Shapes.AddPictureFrame(ShapeType.Rectangle, x, y, width.Value, height.Value, pictureImage);
        else
            slide.Shapes.AddPictureFrame(ShapeType.Rectangle, x, y, pictureImage.Width, pictureImage.Height,
                pictureImage);

        var outputPath = ArgumentHelper.GetAndValidateOutputPath(arguments, path);
        presentation.Save(outputPath, SaveFormat.Pptx);

        return await Task.FromResult($"Image added to slide {slideIndex}: {outputPath}");
    }

    /// <summary>
    ///     Edits image properties
    /// </summary>
    /// <param name="arguments">JSON arguments containing imageIndex, optional x, y, width, height, outputPath</param>
    /// <param name="path">PowerPoint file path</param>
    /// <param name="slideIndex">Slide index (0-based)</param>
    /// <returns>Success message</returns>
    private async Task<string> EditImageAsync(JsonObject? arguments, string path, int slideIndex)
    {
        var shapeIndex = ArgumentHelper.GetInt(arguments, "shapeIndex");
        var imagePath = ArgumentHelper.GetStringNullable(arguments, "imagePath");
        var x = ArgumentHelper.GetFloatNullable(arguments, "x");
        var y = ArgumentHelper.GetFloatNullable(arguments, "y");
        var width = ArgumentHelper.GetFloatNullable(arguments, "width");
        var height = ArgumentHelper.GetFloatNullable(arguments, "height");

        using var presentation = new Presentation(path);
        var slide = PowerPointHelper.GetSlide(presentation, slideIndex);
        var shape = PowerPointHelper.GetShape(slide, shapeIndex);
        if (shape is not PictureFrame pictureFrame)
            throw new ArgumentException($"Shape at index {shapeIndex} is not an image");

        if (!string.IsNullOrEmpty(imagePath) && File.Exists(imagePath))
        {
            await using var imageStream = File.OpenRead(imagePath);
            var newImage = presentation.Images.AddImage(imageStream);
            pictureFrame.PictureFormat.Picture.Image = newImage;
        }

        if (x.HasValue) pictureFrame.X = x.Value;
        if (y.HasValue) pictureFrame.Y = y.Value;
        if (width.HasValue) pictureFrame.Width = width.Value;
        if (height.HasValue) pictureFrame.Height = height.Value;

        var outputPath = ArgumentHelper.GetAndValidateOutputPath(arguments, path);
        presentation.Save(outputPath, SaveFormat.Pptx);
        return await Task.FromResult($"Image updated on slide {slideIndex}, shape {shapeIndex}");
    }
}