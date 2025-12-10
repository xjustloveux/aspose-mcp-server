using System.Text.Json.Nodes;
using Aspose.Slides;
using Aspose.Slides.Export;

namespace AsposeMcpServer.Tools;

public class PptDeleteAudioTool : IAsposeTool
{
    public string Description => "Delete an audio frame from a PowerPoint slide";

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
                description = "Shape index of the audio (0-based)"
            }
        },
        required = new[] { "path", "slideIndex", "shapeIndex" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var slideIndex = arguments?["slideIndex"]?.GetValue<int>() ?? throw new ArgumentException("slideIndex is required");
        var shapeIndex = arguments?["shapeIndex"]?.GetValue<int>() ?? throw new ArgumentException("shapeIndex is required");

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
        if (shape is not IAudioFrame)
        {
            throw new ArgumentException($"Shape at index {shapeIndex} is not an audio frame");
        }

        slide.Shapes.Remove(shape);

        presentation.Save(path, SaveFormat.Pptx);
        return await Task.FromResult($"Audio deleted from slide {slideIndex}, shape {shapeIndex}");
    }
}

