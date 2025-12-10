using System.Text.Json.Nodes;
using Aspose.Slides;
using Aspose.Slides.Animation;
using Aspose.Slides.Export;

namespace AsposeMcpServer.Tools;

public class PptAddAnimationTool : IAsposeTool
{
    public string Description => "Add animation to a shape in PowerPoint";

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
                description = "Shape index (0-based)"
            },
            effectType = new
            {
                type = "string",
                description = "Animation effect type (Fade, Fly, etc.)"
            }
        },
        required = new[] { "path", "slideIndex", "shapeIndex" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var slideIndex = arguments?["slideIndex"]?.GetValue<int>() ?? throw new ArgumentException("slideIndex is required");
        var shapeIndex = arguments?["shapeIndex"]?.GetValue<int>() ?? throw new ArgumentException("shapeIndex is required");
        var effectTypeStr = arguments?["effectType"]?.GetValue<string>() ?? "Fade";

        using var presentation = new Presentation(path);
        var slide = presentation.Slides[slideIndex];
        var shape = slide.Shapes[shapeIndex];

        var effectType = effectTypeStr.ToLower() switch
        {
            "fade" => EffectType.Fade,
            "fly" => EffectType.Fly,
            "appear" => EffectType.Appear,
            "bounce" => EffectType.Bounce,
            "zoom" => EffectType.Zoom,
            _ => EffectType.Fade
        };

        slide.Timeline.MainSequence.AddEffect(shape, effectType, EffectSubtype.None, EffectTriggerType.OnClick);

        presentation.Save(path, SaveFormat.Pptx);

        return await Task.FromResult($"Animation added to shape on slide {slideIndex}: {path}");
    }
}

