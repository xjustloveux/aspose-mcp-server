using System.Text.Json.Nodes;
using Aspose.Slides;
using Aspose.Slides.Export;

namespace AsposeMcpServer.Tools;

public class PptEditImageTool : IAsposeTool
{
    public string Description => "Edit image on a PowerPoint slide (replace or resize)";

    public object InputSchema => new
    {
        type = "object",
        properties = new
        {
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
                description = "Shape index of the image (0-based)"
            },
            imagePath = new
            {
                type = "string",
                description = "New image file path (optional, for replacement)"
            },
            x = new
            {
                type = "number",
                description = "X position (optional)"
            },
            y = new
            {
                type = "number",
                description = "Y position (optional)"
            },
            width = new
            {
                type = "number",
                description = "Width (optional)"
            },
            height = new
            {
                type = "number",
                description = "Height (optional)"
            }
        },
        required = new[] { "path", "slideIndex", "shapeIndex" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var slideIndex = arguments?["slideIndex"]?.GetValue<int>() ?? throw new ArgumentException("slideIndex is required");
        var shapeIndex = arguments?["shapeIndex"]?.GetValue<int>() ?? throw new ArgumentException("shapeIndex is required");
        var imagePath = arguments?["imagePath"]?.GetValue<string>();
        var x = arguments?["x"]?.GetValue<float?>();
        var y = arguments?["y"]?.GetValue<float?>();
        var width = arguments?["width"]?.GetValue<float?>();
        var height = arguments?["height"]?.GetValue<float?>();

        using var presentation = new Presentation(path);
        if (slideIndex < 0 || slideIndex >= presentation.Slides.Count)
        {
            throw new ArgumentException($"slideIndex must be between 0 and {presentation.Slides.Count - 1}");
        }

        var slide = presentation.Slides[slideIndex];
        if (shapeIndex < 0 || shapeIndex >= slide.Shapes.Count)
        {
            throw new ArgumentException($"shapeIndex must be between 0 and {slide.Shapes.Count - 1}");
        }

        var shape = slide.Shapes[shapeIndex];
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

