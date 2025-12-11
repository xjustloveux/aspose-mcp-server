using System.Text.Json.Nodes;
using Aspose.Slides;
using Aspose.Slides.Export;
using AsposeMcpServer.Core;

namespace AsposeMcpServer.Tools;

/// <summary>
/// Unified tool for managing PowerPoint images (add, edit)
/// Merges: PptAddImageTool, PptEditImageTool
/// </summary>
public class PptImageTool : IAsposeTool
{
    public string Description => "Manage PowerPoint images: add or edit";

    public object InputSchema => new
    {
        type = "object",
        properties = new
        {
            operation = new
            {
                type = "string",
                description = "Operation to perform: 'add', 'edit'",
                @enum = new[] { "add", "edit" }
            },
            path = new
            {
                type = "string",
                description = "Presentation file path"
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
            }
        },
        required = new[] { "operation", "path", "slideIndex" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var operation = arguments?["operation"]?.GetValue<string>() ?? throw new ArgumentException("operation is required");
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        SecurityHelper.ValidateFilePath(path, "path");
        var slideIndex = arguments?["slideIndex"]?.GetValue<int>() ?? throw new ArgumentException("slideIndex is required");

        return operation.ToLower() switch
        {
            "add" => await AddImageAsync(arguments, path, slideIndex),
            "edit" => await EditImageAsync(arguments, path, slideIndex),
            _ => throw new ArgumentException($"Unknown operation: {operation}")
        };
    }

    private async Task<string> AddImageAsync(JsonObject? arguments, string path, int slideIndex)
    {
        var imagePath = arguments?["imagePath"]?.GetValue<string>() ?? throw new ArgumentException("imagePath is required for add operation");
        var x = arguments?["x"]?.GetValue<float>() ?? 100;
        var y = arguments?["y"]?.GetValue<float>() ?? 100;
        var width = arguments?["width"]?.GetValue<float>();
        var height = arguments?["height"]?.GetValue<float>();

        if (!File.Exists(imagePath))
        {
            throw new FileNotFoundException($"Image file not found: {imagePath}");
        }

        using var presentation = new Presentation(path);
        if (slideIndex < 0 || slideIndex >= presentation.Slides.Count)
        {
            throw new ArgumentException($"slideIndex must be between 0 and {presentation.Slides.Count - 1}");
        }

        var slide = presentation.Slides[slideIndex];

        using var image = File.OpenRead(imagePath);
        var pictureImage = presentation.Images.AddImage(image);

        if (width.HasValue && height.HasValue)
        {
            slide.Shapes.AddPictureFrame(ShapeType.Rectangle, x, y, width.Value, height.Value, pictureImage);
        }
        else
        {
            slide.Shapes.AddPictureFrame(ShapeType.Rectangle, x, y, pictureImage.Width, pictureImage.Height, pictureImage);
        }

        presentation.Save(path, SaveFormat.Pptx);

        return await Task.FromResult($"Image added to slide {slideIndex}: {path}");
    }

    private async Task<string> EditImageAsync(JsonObject? arguments, string path, int slideIndex)
    {
        var shapeIndex = arguments?["shapeIndex"]?.GetValue<int>() ?? throw new ArgumentException("shapeIndex is required for edit operation");
        var imagePath = arguments?["imagePath"]?.GetValue<string>();
        var x = arguments?["x"]?.GetValue<float?>();
        var y = arguments?["y"]?.GetValue<float?>();
        var width = arguments?["width"]?.GetValue<float?>();
        var height = arguments?["height"]?.GetValue<float?>();

        using var presentation = new Presentation(path);
        var slide = PowerPointHelper.GetSlide(presentation, slideIndex);
        var shape = PowerPointHelper.GetShape(slide, shapeIndex);
        if (shape is not PictureFrame pictureFrame)
        {
            throw new ArgumentException($"Shape at index {shapeIndex} is not an image");
        }

        if (!string.IsNullOrEmpty(imagePath) && File.Exists(imagePath))
        {
            using var imageStream = File.OpenRead(imagePath);
            var newImage = presentation.Images.AddImage(imageStream);
            pictureFrame.PictureFormat.Picture.Image = newImage;
        }

        if (x.HasValue) pictureFrame.X = x.Value;
        if (y.HasValue) pictureFrame.Y = y.Value;
        if (width.HasValue) pictureFrame.Width = width.Value;
        if (height.HasValue) pictureFrame.Height = height.Value;

        presentation.Save(path, SaveFormat.Pptx);
        return await Task.FromResult($"Image updated on slide {slideIndex}, shape {shapeIndex}");
    }
}

