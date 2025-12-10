using System.Text.Json.Nodes;
using Aspose.Slides;
using Aspose.Slides.Export;

namespace AsposeMcpServer.Tools;

public class PptAddImageTool : IAsposeTool
{
    public string Description => "Add an image to a PowerPoint slide";

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
            imagePath = new
            {
                type = "string",
                description = "Image file path"
            },
            x = new
            {
                type = "number",
                description = "X position (optional, default: 100)"
            },
            y = new
            {
                type = "number",
                description = "Y position (optional, default: 100)"
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
        required = new[] { "path", "slideIndex", "imagePath" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var slideIndex = arguments?["slideIndex"]?.GetValue<int>() ?? throw new ArgumentException("slideIndex is required");
        var imagePath = arguments?["imagePath"]?.GetValue<string>() ?? throw new ArgumentException("imagePath is required");
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
}

