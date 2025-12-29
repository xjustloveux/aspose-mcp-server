using System.Text.Json.Nodes;
using Aspose.Slides;
using Aspose.Slides.Animation;
using Aspose.Slides.Export;
using AsposeMcpServer.Core;

namespace AsposeMcpServer.Tools.PowerPoint;

/// <summary>
///     Unified tool for managing PowerPoint animations (add, edit, delete).
/// </summary>
public class PptAnimationTool : IAsposeTool
{
    /// <inheritdoc />
    public string Description => @"Manage PowerPoint animations. Supports 3 operations: add, edit, delete.

Usage examples:
- Add animation: ppt_animation(operation='add', path='presentation.pptx', slideIndex=0, shapeIndex=0, effectType='Fade')
- Edit animation: ppt_animation(operation='edit', path='presentation.pptx', slideIndex=0, shapeIndex=0, animationIndex=0, effectType='Fly')
- Delete animation: ppt_animation(operation='delete', path='presentation.pptx', slideIndex=0, shapeIndex=0)";

    /// <inheritdoc />
    public object InputSchema => new
    {
        type = "object",
        properties = new
        {
            operation = new
            {
                type = "string",
                description = @"Operation to perform.
- 'add': Add animation to shape
- 'edit': Edit existing animation (use animationIndex to specify which one)
- 'delete': Delete animation(s)",
                @enum = new[] { "add", "edit", "delete" }
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
                description = "Shape index (0-based, required for add/edit, optional for delete)"
            },
            animationIndex = new
            {
                type = "number",
                description = "Animation index (0-based, optional for edit/delete, targets specific animation)"
            },
            effectType = new
            {
                type = "string",
                description = "Animation effect type (e.g., Fade, Fly, Appear, Bounce, Zoom, Wipe, Split, etc.)"
            },
            effectSubtype = new
            {
                type = "string",
                description =
                    "Animation effect subtype for direction/style (e.g., FromBottom, FromLeft, FromRight, FromTop, Horizontal, Vertical)"
            },
            triggerType = new
            {
                type = "string",
                description = "Trigger type (OnClick, AfterPrevious, WithPrevious)"
            },
            duration = new
            {
                type = "number",
                description = "Animation duration in seconds"
            },
            delay = new
            {
                type = "number",
                description = "Animation delay in seconds"
            },
            outputPath = new
            {
                type = "string",
                description = "Output file path (optional, defaults to input path)"
            }
        },
        required = new[] { "operation", "path", "slideIndex" }
    };

    /// <inheritdoc />
    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var operation = ArgumentHelper.GetString(arguments, "operation");
        var path = ArgumentHelper.GetAndValidatePath(arguments);
        var outputPath = ArgumentHelper.GetAndValidateOutputPath(arguments, path);
        var slideIndex = ArgumentHelper.GetInt(arguments, "slideIndex");

        return operation.ToLower() switch
        {
            "add" => await AddAnimationAsync(path, outputPath, slideIndex, arguments),
            "edit" => await EditAnimationAsync(path, outputPath, slideIndex, arguments),
            "delete" => await DeleteAnimationAsync(path, outputPath, slideIndex, arguments),
            _ => throw new ArgumentException($"Unknown operation: {operation}")
        };
    }

    /// <summary>
    ///     Adds animation to a shape.
    /// </summary>
    /// <param name="path">PowerPoint file path.</param>
    /// <param name="outputPath">Output file path.</param>
    /// <param name="slideIndex">Slide index (0-based).</param>
    /// <param name="arguments">JSON arguments containing shapeIndex, effectType, effectSubtype, triggerType.</param>
    /// <returns>Success message.</returns>
    /// <exception cref="ArgumentException">Thrown when shapeIndex is out of range.</exception>
    private Task<string> AddAnimationAsync(string path, string outputPath, int slideIndex, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var shapeIndex = ArgumentHelper.GetInt(arguments, "shapeIndex");
            var effectTypeStr = ArgumentHelper.GetString(arguments, "effectType", "Fade");
            var effectSubtypeStr = ArgumentHelper.GetStringNullable(arguments, "effectSubtype");
            var triggerTypeStr = ArgumentHelper.GetStringNullable(arguments, "triggerType");

            using var presentation = new Presentation(path);
            var slide = PowerPointHelper.GetSlide(presentation, slideIndex);
            PowerPointHelper.ValidateShapeIndex(shapeIndex, slide);
            var shape = slide.Shapes[shapeIndex];

            var effectType = ParseEffectType(effectTypeStr);
            var effectSubtype = ParseEffectSubtype(effectSubtypeStr);
            var triggerType = ParseTriggerType(triggerTypeStr);

            slide.Timeline.MainSequence.AddEffect(shape, effectType, effectSubtype, triggerType);
            presentation.Save(outputPath, SaveFormat.Pptx);

            return
                $"Animation '{effectTypeStr}' added to shape {shapeIndex} on slide {slideIndex}. Output: {outputPath}";
        });
    }

    /// <summary>
    ///     Edits animation properties. If animationIndex is provided, modifies that specific animation;
    ///     otherwise replaces all animations for the shape.
    /// </summary>
    /// <param name="path">PowerPoint file path.</param>
    /// <param name="outputPath">Output file path.</param>
    /// <param name="slideIndex">Slide index (0-based).</param>
    /// <param name="arguments">
    ///     JSON arguments containing shapeIndex, animationIndex, effectType, effectSubtype, triggerType,
    ///     duration, delay.
    /// </param>
    /// <returns>Success message.</returns>
    /// <exception cref="ArgumentException">Thrown when shapeIndex or animationIndex is out of range.</exception>
    private Task<string> EditAnimationAsync(string path, string outputPath, int slideIndex, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var shapeIndex = ArgumentHelper.GetInt(arguments, "shapeIndex");
            var animationIndex = ArgumentHelper.GetIntNullable(arguments, "animationIndex");
            var effectTypeStr = ArgumentHelper.GetStringNullable(arguments, "effectType");
            var effectSubtypeStr = ArgumentHelper.GetStringNullable(arguments, "effectSubtype");
            var triggerTypeStr = ArgumentHelper.GetStringNullable(arguments, "triggerType");
            var duration = ArgumentHelper.GetFloatNullable(arguments, "duration");
            var delay = ArgumentHelper.GetFloatNullable(arguments, "delay");

            using var presentation = new Presentation(path);
            var slide = PowerPointHelper.GetSlide(presentation, slideIndex);
            PowerPointHelper.ValidateShapeIndex(shapeIndex, slide);

            var shape = slide.Shapes[shapeIndex];
            var sequence = slide.Timeline.MainSequence;
            var animations = sequence.Where(e => e.TargetShape == shape).ToList();

            if (animationIndex.HasValue)
            {
                if (animationIndex.Value < 0 || animationIndex.Value >= animations.Count)
                    throw new ArgumentException($"animationIndex must be between 0 and {animations.Count - 1}");

                var effect = animations[animationIndex.Value];
                if (duration.HasValue) effect.Timing.Duration = duration.Value;
                if (delay.HasValue) effect.Timing.TriggerDelayTime = delay.Value;

                if (!string.IsNullOrEmpty(effectTypeStr) || !string.IsNullOrEmpty(effectSubtypeStr) ||
                    !string.IsNullOrEmpty(triggerTypeStr))
                {
                    var newEffectType = !string.IsNullOrEmpty(effectTypeStr)
                        ? ParseEffectType(effectTypeStr)
                        : effect.Type;
                    var newSubtype = !string.IsNullOrEmpty(effectSubtypeStr)
                        ? ParseEffectSubtype(effectSubtypeStr)
                        : effect.Subtype;
                    var newTrigger = !string.IsNullOrEmpty(triggerTypeStr)
                        ? ParseTriggerType(triggerTypeStr)
                        : effect.Timing.TriggerType;

                    var currentDuration = effect.Timing.Duration;
                    var currentDelay = effect.Timing.TriggerDelayTime;

                    sequence.Remove(effect);
                    var newEffect = sequence.AddEffect(shape, newEffectType, newSubtype, newTrigger);
                    newEffect.Timing.Duration = duration ?? currentDuration;
                    newEffect.Timing.TriggerDelayTime = delay ?? currentDelay;
                }
            }
            else
            {
                foreach (var anim in animations)
                    sequence.Remove(anim);

                if (!string.IsNullOrEmpty(effectTypeStr))
                {
                    var effectType = ParseEffectType(effectTypeStr);
                    var effectSubtype = ParseEffectSubtype(effectSubtypeStr);
                    var triggerType = ParseTriggerType(triggerTypeStr);
                    var effect = sequence.AddEffect(shape, effectType, effectSubtype, triggerType);

                    if (duration.HasValue) effect.Timing.Duration = duration.Value;
                    if (delay.HasValue) effect.Timing.TriggerDelayTime = delay.Value;
                }
            }

            presentation.Save(outputPath, SaveFormat.Pptx);
            return $"Animation updated on slide {slideIndex}, shape {shapeIndex}. Output: {outputPath}";
        });
    }

    /// <summary>
    ///     Deletes animation(s) from a shape or slide.
    /// </summary>
    /// <param name="path">PowerPoint file path.</param>
    /// <param name="outputPath">Output file path.</param>
    /// <param name="slideIndex">Slide index (0-based).</param>
    /// <param name="arguments">JSON arguments containing optional shapeIndex and animationIndex.</param>
    /// <returns>Success message.</returns>
    /// <exception cref="ArgumentException">Thrown when shapeIndex or animationIndex is out of range.</exception>
    private Task<string> DeleteAnimationAsync(string path, string outputPath, int slideIndex, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var shapeIndex = ArgumentHelper.GetIntNullable(arguments, "shapeIndex");
            var animationIndex = ArgumentHelper.GetIntNullable(arguments, "animationIndex");

            using var presentation = new Presentation(path);
            var slide = PowerPointHelper.GetSlide(presentation, slideIndex);
            var sequence = slide.Timeline.MainSequence;

            if (shapeIndex.HasValue)
            {
                PowerPointHelper.ValidateShapeIndex(shapeIndex.Value, slide);
                var shape = slide.Shapes[shapeIndex.Value];
                var animations = sequence.Where(e => e.TargetShape == shape).ToList();

                if (animationIndex.HasValue)
                {
                    if (animationIndex.Value < 0 || animationIndex.Value >= animations.Count)
                        throw new ArgumentException($"animationIndex must be between 0 and {animations.Count - 1}");
                    sequence.Remove(animations[animationIndex.Value]);
                }
                else
                {
                    foreach (var anim in animations)
                        sequence.Remove(anim);
                }
            }
            else
            {
                sequence.Clear();
            }

            presentation.Save(outputPath, SaveFormat.Pptx);
            return $"Animation(s) deleted from slide {slideIndex}. Output: {outputPath}";
        });
    }

    /// <summary>
    ///     Parses effect type string to EffectType enum.
    /// </summary>
    /// <param name="effectTypeStr">Effect type string.</param>
    /// <returns>Parsed EffectType, defaults to Fade if invalid.</returns>
    private static EffectType ParseEffectType(string? effectTypeStr)
    {
        if (string.IsNullOrEmpty(effectTypeStr)) return EffectType.Fade;
        return Enum.TryParse<EffectType>(effectTypeStr, true, out var result) ? result : EffectType.Fade;
    }

    /// <summary>
    ///     Parses effect subtype string to EffectSubtype enum.
    /// </summary>
    /// <param name="subtypeStr">Effect subtype string.</param>
    /// <returns>Parsed EffectSubtype, defaults to None if invalid.</returns>
    private static EffectSubtype ParseEffectSubtype(string? subtypeStr)
    {
        if (string.IsNullOrEmpty(subtypeStr)) return EffectSubtype.None;
        return Enum.TryParse<EffectSubtype>(subtypeStr, true, out var result) ? result : EffectSubtype.None;
    }

    /// <summary>
    ///     Parses trigger type string to EffectTriggerType enum.
    /// </summary>
    /// <param name="triggerTypeStr">Trigger type string.</param>
    /// <returns>Parsed EffectTriggerType, defaults to OnClick if invalid.</returns>
    private static EffectTriggerType ParseTriggerType(string? triggerTypeStr)
    {
        if (string.IsNullOrEmpty(triggerTypeStr)) return EffectTriggerType.OnClick;
        return Enum.TryParse<EffectTriggerType>(triggerTypeStr, true, out var result)
            ? result
            : EffectTriggerType.OnClick;
    }
}