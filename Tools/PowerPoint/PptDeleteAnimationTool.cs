using System.Text.Json.Nodes;
using Aspose.Slides;
using Aspose.Slides.Animation;
using Aspose.Slides.Export;

namespace AsposeMcpServer.Tools;

public class PptDeleteAnimationTool : IAsposeTool
{
    public string Description => "Delete animation(s) from a shape or all animations from a slide";

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
                description = "Shape index (0-based, optional, if not provided deletes all animations from slide)"
            },
            animationIndex = new
            {
                type = "number",
                description = "Animation index (0-based, optional, if not provided deletes all animations for the shape)"
            }
        },
        required = new[] { "path", "slideIndex" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var slideIndex = arguments?["slideIndex"]?.GetValue<int>() ?? throw new ArgumentException("slideIndex is required");
        var shapeIndex = arguments?["shapeIndex"]?.GetValue<int?>();
        var animationIndex = arguments?["animationIndex"]?.GetValue<int?>();

        using var presentation = new Presentation(path);
        if (slideIndex < 0 || slideIndex >= presentation.Slides.Count)
        {
            throw new ArgumentException($"slideIndex must be between 0 and {presentation.Slides.Count - 1}");
        }

        var slide = presentation.Slides[slideIndex];
        var sequence = slide.Timeline.MainSequence;

        if (shapeIndex.HasValue)
        {
            if (shapeIndex.Value < 0 || shapeIndex.Value >= slide.Shapes.Count)
            {
                throw new ArgumentException($"shapeIndex must be between 0 and {slide.Shapes.Count - 1}");
            }

            var shape = slide.Shapes[shapeIndex.Value];
            var animations = sequence.Where(e => e.TargetShape == shape).ToList();

            if (animationIndex.HasValue)
            {
                if (animationIndex.Value < 0 || animationIndex.Value >= animations.Count)
                {
                    throw new ArgumentException($"animationIndex must be between 0 and {animations.Count - 1}");
                }
                sequence.Remove(animations[animationIndex.Value]);
            }
            else
            {
                foreach (var anim in animations)
                {
                    sequence.Remove(anim);
                }
            }
        }
        else
        {
            // Delete all animations from the slide
            sequence.Clear();
        }

        presentation.Save(path, SaveFormat.Pptx);
        return await Task.FromResult($"Animation(s) deleted from slide {slideIndex}");
    }
}

