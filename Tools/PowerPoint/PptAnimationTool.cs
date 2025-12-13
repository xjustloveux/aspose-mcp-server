using System.Text.Json.Nodes;
using System.Linq;
using Aspose.Slides;
using Aspose.Slides.Animation;
using Aspose.Slides.Export;
using AsposeMcpServer.Core;

namespace AsposeMcpServer.Tools;

/// <summary>
/// Unified tool for managing PowerPoint animations (add, edit, delete)
/// Merges: PptAddAnimationTool, PptEditAnimationTool, PptDeleteAnimationTool
/// </summary>
public class PptAnimationTool : IAsposeTool
{
    public string Description => @"Manage PowerPoint animations. Supports 3 operations: add, edit, delete.

Usage examples:
- Add animation: ppt_animation(operation='add', path='presentation.pptx', slideIndex=0, shapeIndex=0, effectType='Fade')
- Edit animation: ppt_animation(operation='edit', path='presentation.pptx', slideIndex=0, shapeIndex=0, effectType='Fly')
- Delete animation: ppt_animation(operation='delete', path='presentation.pptx', slideIndex=0, shapeIndex=0)";

    public object InputSchema => new
    {
        type = "object",
        properties = new
        {
            operation = new
            {
                type = "string",
                description = @"Operation to perform.
- 'add': Add animation to shape (required params: path, slideIndex, shapeIndex, effectType)
- 'edit': Edit animation (required params: path, slideIndex, shapeIndex)
- 'delete': Delete animation (required params: path, slideIndex, shapeIndex)",
                @enum = new[] { "add", "edit", "delete" }
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
                description = "Shape index (0-based, required for add/edit, optional for delete)"
            },
            effectType = new
            {
                type = "string",
                description = "Animation effect type (Fade, Fly, Appear, Bounce, Zoom, etc., required for add, optional for edit)"
            },
            triggerType = new
            {
                type = "string",
                description = "Trigger type (OnClick, AfterPrevious, WithPrevious, optional, for edit)"
            },
            duration = new
            {
                type = "number",
                description = "Animation duration in seconds (optional, for edit)"
            },
            delay = new
            {
                type = "number",
                description = "Animation delay in seconds (optional, for edit)"
            },
            animationIndex = new
            {
                type = "number",
                description = "Animation index (0-based, optional, for delete, if not provided deletes all animations for the shape)"
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
            "add" => await AddAnimationAsync(arguments, path, slideIndex),
            "edit" => await EditAnimationAsync(arguments, path, slideIndex),
            "delete" => await DeleteAnimationAsync(arguments, path, slideIndex),
            _ => throw new ArgumentException($"Unknown operation: {operation}")
        };
    }

    private async Task<string> AddAnimationAsync(JsonObject? arguments, string path, int slideIndex)
    {
        var shapeIndex = arguments?["shapeIndex"]?.GetValue<int>() ?? throw new ArgumentException("shapeIndex is required for add operation");
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

    private async Task<string> EditAnimationAsync(JsonObject? arguments, string path, int slideIndex)
    {
        var shapeIndex = arguments?["shapeIndex"]?.GetValue<int>() ?? throw new ArgumentException("shapeIndex is required for edit operation");
        var effectTypeStr = arguments?["effectType"]?.GetValue<string>();
        var triggerTypeStr = arguments?["triggerType"]?.GetValue<string>();
        var duration = arguments?["duration"]?.GetValue<float?>();
        var delay = arguments?["delay"]?.GetValue<float?>();

        using var presentation = new Presentation(path);
        var slide = PowerPointHelper.GetSlide(presentation, slideIndex);
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

    private async Task<string> DeleteAnimationAsync(JsonObject? arguments, string path, int slideIndex)
    {
        var shapeIndex = arguments?["shapeIndex"]?.GetValue<int?>();
        var animationIndex = arguments?["animationIndex"]?.GetValue<int?>();

        using var presentation = new Presentation(path);
        var slide = PowerPointHelper.GetSlide(presentation, slideIndex);
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

