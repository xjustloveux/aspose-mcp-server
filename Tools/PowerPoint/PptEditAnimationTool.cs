using System.Text.Json.Nodes;
using Aspose.Slides;
using Aspose.Slides.Animation;
using Aspose.Slides.Export;

namespace AsposeMcpServer.Tools;

public class PptEditAnimationTool : IAsposeTool
{
    public string Description => "Edit animation effects on a shape in PowerPoint";

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
                description = "Animation effect type (Fade, Fly, Appear, Bounce, Zoom, etc., optional)"
            },
            triggerType = new
            {
                type = "string",
                description = "Trigger type (OnClick, AfterPrevious, WithPrevious, optional)"
            },
            duration = new
            {
                type = "number",
                description = "Animation duration in seconds (optional)"
            },
            delay = new
            {
                type = "number",
                description = "Animation delay in seconds (optional)"
            }
        },
        required = new[] { "path", "slideIndex", "shapeIndex" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var slideIndex = arguments?["slideIndex"]?.GetValue<int>() ?? throw new ArgumentException("slideIndex is required");
        var shapeIndex = arguments?["shapeIndex"]?.GetValue<int>() ?? throw new ArgumentException("shapeIndex is required");
        var effectTypeStr = arguments?["effectType"]?.GetValue<string>();
        var triggerTypeStr = arguments?["triggerType"]?.GetValue<string>();
        var duration = arguments?["duration"]?.GetValue<float?>();
        var delay = arguments?["delay"]?.GetValue<float?>();

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
        var sequence = slide.Timeline.MainSequence;

        // Remove existing animations for this shape
        for (int i = sequence.Count - 1; i >= 0; i--)
        {
            if (sequence[i].TargetShape == shape)
            {
                sequence.Remove(sequence[i]);
            }
        }

        // Add new animation if specified
        if (!string.IsNullOrEmpty(effectTypeStr))
        {
            var effectType = effectTypeStr.ToLower() switch
            {
                "fade" => EffectType.Fade,
                "fly" => EffectType.Fly,
                "appear" => EffectType.Appear,
                "bounce" => EffectType.Bounce,
                "zoom" => EffectType.Zoom,
                "wipe" => EffectType.Wipe,
                "dissolve" => EffectType.Dissolve,
                "randombars" => EffectType.RandomBars,
                "randomeffects" => EffectType.RandomEffects,
                _ => EffectType.Fade
            };

            var triggerType = triggerTypeStr?.ToLower() switch
            {
                "afterprevious" => EffectTriggerType.AfterPrevious,
                "withprevious" => EffectTriggerType.WithPrevious,
                "onclick" => EffectTriggerType.OnClick,
                _ => EffectTriggerType.OnClick
            };

            var effect = sequence.AddEffect(shape, effectType, EffectSubtype.None, triggerType);

            if (duration.HasValue)
            {
                effect.Timing.Duration = duration.Value;
            }

            if (delay.HasValue)
            {
                effect.Timing.TriggerDelayTime = delay.Value;
            }
        }

        presentation.Save(path, SaveFormat.Pptx);
        return await Task.FromResult($"Animation updated on slide {slideIndex}, shape {shapeIndex}");
    }
}

