using System.ComponentModel;
using System.Text.Json;
using Aspose.Slides;
using Aspose.Slides.Animation;
using AsposeMcpServer.Core.Helpers;
using AsposeMcpServer.Core.Session;
using ModelContextProtocol.Server;

namespace AsposeMcpServer.Tools.PowerPoint;

/// <summary>
///     Unified tool for managing PowerPoint animations (add, edit, delete, get).
/// </summary>
[McpServerToolType]
public class PptAnimationTool
{
    /// <summary>
    ///     Identity accessor for session isolation.
    /// </summary>
    private readonly ISessionIdentityAccessor? _identityAccessor;

    /// <summary>
    ///     Session manager for document lifecycle management.
    /// </summary>
    private readonly DocumentSessionManager? _sessionManager;

    /// <summary>
    ///     Initializes a new instance of the <see cref="PptAnimationTool" /> class.
    /// </summary>
    /// <param name="sessionManager">Optional session manager for in-memory document editing.</param>
    /// <param name="identityAccessor">Optional identity accessor for session isolation.</param>
    public PptAnimationTool(DocumentSessionManager? sessionManager = null,
        ISessionIdentityAccessor? identityAccessor = null)
    {
        _sessionManager = sessionManager;
        _identityAccessor = identityAccessor;
    }

    /// <summary>
    ///     Executes a PowerPoint animation operation (add, edit, delete, get).
    /// </summary>
    /// <param name="operation">The operation to perform: add, edit, delete, get.</param>
    /// <param name="slideIndex">Slide index (0-based).</param>
    /// <param name="path">Presentation file path (required if no sessionId).</param>
    /// <param name="sessionId">Session ID for in-memory editing.</param>
    /// <param name="outputPath">Output file path (file mode only).</param>
    /// <param name="shapeIndex">Shape index (0-based, required for add/edit, optional for delete).</param>
    /// <param name="animationIndex">Animation index (0-based, optional for edit/delete, targets specific animation).</param>
    /// <param name="effectType">Animation effect type (e.g., Fade, Fly, Appear, Bounce, Zoom, Wipe, Split, etc.).</param>
    /// <param name="effectSubtype">Animation effect subtype for direction/style.</param>
    /// <param name="triggerType">Trigger type (OnClick, AfterPrevious, WithPrevious).</param>
    /// <param name="duration">Animation duration in seconds.</param>
    /// <param name="delay">Animation delay in seconds.</param>
    /// <returns>A message indicating the result of the operation, or JSON data for get operations.</returns>
    /// <exception cref="ArgumentException">Thrown when required parameters are missing or the operation is unknown.</exception>
    [McpServerTool(Name = "ppt_animation")]
    [Description(@"Manage PowerPoint animations. Supports 4 operations: add, edit, delete, get.

Usage examples:
- Add animation: ppt_animation(operation='add', path='presentation.pptx', slideIndex=0, shapeIndex=0, effectType='Fade')
- Edit animation: ppt_animation(operation='edit', path='presentation.pptx', slideIndex=0, shapeIndex=0, animationIndex=0, effectType='Fly')
- Delete animation: ppt_animation(operation='delete', path='presentation.pptx', slideIndex=0, shapeIndex=0)
- Get animations: ppt_animation(operation='get', path='presentation.pptx', slideIndex=0)")]
    public string Execute(
        [Description("Operation: add, edit, delete, get")]
        string operation,
        [Description("Slide index (0-based)")] int slideIndex,
        [Description("Presentation file path (required if no sessionId)")]
        string? path = null,
        [Description("Session ID for in-memory editing")]
        string? sessionId = null,
        [Description("Output file path (file mode only)")]
        string? outputPath = null,
        [Description("Shape index (0-based, required for add/edit, optional for delete)")]
        int? shapeIndex = null,
        [Description("Animation index (0-based, optional for edit/delete, targets specific animation)")]
        int? animationIndex = null,
        [Description("Animation effect type (e.g., Fade, Fly, Appear, Bounce, Zoom, Wipe, Split, etc.)")]
        string? effectType = null,
        [Description(
            "Animation effect subtype for direction/style (e.g., FromBottom, FromLeft, FromRight, FromTop, Horizontal, Vertical)")]
        string? effectSubtype = null,
        [Description("Trigger type (OnClick, AfterPrevious, WithPrevious)")]
        string? triggerType = null,
        [Description("Animation duration in seconds")]
        float? duration = null,
        [Description("Animation delay in seconds")]
        float? delay = null)
    {
        using var ctx = DocumentContext<Presentation>.Create(_sessionManager, sessionId, path, _identityAccessor);

        return operation.ToLower() switch
        {
            "add" => AddAnimation(ctx, outputPath, slideIndex, shapeIndex, effectType, effectSubtype, triggerType),
            "edit" => EditAnimation(ctx, outputPath, slideIndex, shapeIndex, animationIndex, effectType, effectSubtype,
                triggerType, duration, delay),
            "delete" => DeleteAnimation(ctx, outputPath, slideIndex, shapeIndex, animationIndex),
            "get" => GetAnimations(ctx, slideIndex, shapeIndex),
            _ => throw new ArgumentException($"Unknown operation: {operation}")
        };
    }

    /// <summary>
    ///     Gets animation information for a slide or specific shape.
    /// </summary>
    /// <param name="ctx">The document context.</param>
    /// <param name="slideIndex">The slide index (0-based).</param>
    /// <param name="shapeIndex">The shape index (0-based, optional). If provided, only returns animations for that shape.</param>
    /// <returns>A JSON string containing the animation information.</returns>
    private static string GetAnimations(DocumentContext<Presentation> ctx, int slideIndex, int? shapeIndex)
    {
        var presentation = ctx.Document;
        var slide = PowerPointHelper.GetSlide(presentation, slideIndex);
        var sequence = slide.Timeline.MainSequence;

        List<object> animations = [];
        var index = 0;

        foreach (var effect in sequence)
        {
            // If shapeIndex is specified, filter by that shape
            if (shapeIndex.HasValue)
            {
                var targetShapeIndex = slide.Shapes.IndexOf(effect.TargetShape);
                if (targetShapeIndex != shapeIndex.Value)
                    continue;
            }

            var shapeName = effect.TargetShape?.Name ?? "(unknown)";
            var shapeIdx = effect.TargetShape != null ? slide.Shapes.IndexOf(effect.TargetShape) : -1;

            animations.Add(new
            {
                index,
                shapeIndex = shapeIdx,
                shapeName,
                effectType = effect.Type.ToString(),
                effectSubtype = effect.Subtype.ToString(),
                triggerType = effect.Timing.TriggerType.ToString(),
                duration = effect.Timing.Duration,
                delay = effect.Timing.TriggerDelayTime
            });

            index++;
        }

        var result = new
        {
            slideIndex,
            filterByShapeIndex = shapeIndex,
            totalAnimationsOnSlide = sequence.Count,
            animations
        };

        return JsonSerializer.Serialize(result, new JsonSerializerOptions { WriteIndented = true });
    }

    /// <summary>
    ///     Adds animation to a shape.
    /// </summary>
    /// <param name="ctx">The document context.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="slideIndex">The slide index (0-based).</param>
    /// <param name="shapeIndex">The shape index (0-based).</param>
    /// <param name="effectTypeStr">The animation effect type string.</param>
    /// <param name="effectSubtypeStr">The animation effect subtype string.</param>
    /// <param name="triggerTypeStr">The animation trigger type string.</param>
    /// <returns>A message indicating the result of the operation.</returns>
    /// <exception cref="ArgumentException">Thrown when shapeIndex is not provided.</exception>
    private static string AddAnimation(DocumentContext<Presentation> ctx, string? outputPath, int slideIndex,
        int? shapeIndex, string? effectTypeStr, string? effectSubtypeStr, string? triggerTypeStr)
    {
        if (!shapeIndex.HasValue)
            throw new ArgumentException("shapeIndex is required for add operation");

        var presentation = ctx.Document;
        var slide = PowerPointHelper.GetSlide(presentation, slideIndex);
        PowerPointHelper.ValidateShapeIndex(shapeIndex.Value, slide);
        var shape = slide.Shapes[shapeIndex.Value];

        var effectType = ParseEffectType(effectTypeStr ?? "Fade");
        var effectSubtype = ParseEffectSubtype(effectSubtypeStr);
        var triggerType = ParseTriggerType(triggerTypeStr);

        slide.Timeline.MainSequence.AddEffect(shape, effectType, effectSubtype, triggerType);

        ctx.Save(outputPath);

        var result = $"Animation '{effectTypeStr ?? "Fade"}' added to shape {shapeIndex} on slide {slideIndex}.\n";
        result += ctx.GetOutputMessage(outputPath);
        return result;
    }

    /// <summary>
    ///     Edits animation properties. If animationIndex is provided, modifies that specific animation;
    ///     otherwise replaces all animations for the shape.
    /// </summary>
    /// <param name="ctx">The document context.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="slideIndex">The slide index (0-based).</param>
    /// <param name="shapeIndex">The shape index (0-based).</param>
    /// <param name="animationIndex">The animation index (0-based, optional).</param>
    /// <param name="effectTypeStr">The animation effect type string.</param>
    /// <param name="effectSubtypeStr">The animation effect subtype string.</param>
    /// <param name="triggerTypeStr">The animation trigger type string.</param>
    /// <param name="duration">The animation duration in seconds.</param>
    /// <param name="delay">The animation delay in seconds.</param>
    /// <returns>A message indicating the result of the operation.</returns>
    /// <exception cref="ArgumentException">Thrown when shapeIndex is not provided or animationIndex is out of range.</exception>
    private static string EditAnimation(DocumentContext<Presentation> ctx, string? outputPath, int slideIndex,
        int? shapeIndex, int? animationIndex, string? effectTypeStr, string? effectSubtypeStr,
        string? triggerTypeStr, float? duration, float? delay)
    {
        if (!shapeIndex.HasValue)
            throw new ArgumentException("shapeIndex is required for edit operation");

        var presentation = ctx.Document;
        var slide = PowerPointHelper.GetSlide(presentation, slideIndex);
        PowerPointHelper.ValidateShapeIndex(shapeIndex.Value, slide);

        var shape = slide.Shapes[shapeIndex.Value];
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

        ctx.Save(outputPath);

        var result = $"Animation updated on slide {slideIndex}, shape {shapeIndex}.\n";
        result += ctx.GetOutputMessage(outputPath);
        return result;
    }

    /// <summary>
    ///     Deletes animation(s) from a shape or slide.
    /// </summary>
    /// <param name="ctx">The document context.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="slideIndex">The slide index (0-based).</param>
    /// <param name="shapeIndex">The shape index (0-based, optional).</param>
    /// <param name="animationIndex">The animation index (0-based, optional).</param>
    /// <returns>A message indicating the result of the operation.</returns>
    /// <exception cref="ArgumentException">Thrown when animationIndex is out of range.</exception>
    private static string DeleteAnimation(DocumentContext<Presentation> ctx, string? outputPath, int slideIndex,
        int? shapeIndex, int? animationIndex)
    {
        var presentation = ctx.Document;
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

        ctx.Save(outputPath);

        var result = $"Animation(s) deleted from slide {slideIndex}.\n";
        result += ctx.GetOutputMessage(outputPath);
        return result;
    }

    /// <summary>
    ///     Parses effect type string to EffectType enum.
    /// </summary>
    /// <param name="effectTypeStr">The effect type string to parse.</param>
    /// <returns>The parsed EffectType enum value, or Fade as default.</returns>
    private static EffectType ParseEffectType(string? effectTypeStr)
    {
        if (string.IsNullOrEmpty(effectTypeStr)) return EffectType.Fade;
        return Enum.TryParse<EffectType>(effectTypeStr, true, out var result) ? result : EffectType.Fade;
    }

    /// <summary>
    ///     Parses effect subtype string to EffectSubtype enum.
    /// </summary>
    /// <param name="subtypeStr">The effect subtype string to parse.</param>
    /// <returns>The parsed EffectSubtype enum value, or None as default.</returns>
    private static EffectSubtype ParseEffectSubtype(string? subtypeStr)
    {
        if (string.IsNullOrEmpty(subtypeStr)) return EffectSubtype.None;
        return Enum.TryParse<EffectSubtype>(subtypeStr, true, out var result) ? result : EffectSubtype.None;
    }

    /// <summary>
    ///     Parses trigger type string to EffectTriggerType enum.
    /// </summary>
    /// <param name="triggerTypeStr">The trigger type string to parse.</param>
    /// <returns>The parsed EffectTriggerType enum value, or OnClick as default.</returns>
    private static EffectTriggerType ParseTriggerType(string? triggerTypeStr)
    {
        if (string.IsNullOrEmpty(triggerTypeStr)) return EffectTriggerType.OnClick;
        return Enum.TryParse<EffectTriggerType>(triggerTypeStr, true, out var result)
            ? result
            : EffectTriggerType.OnClick;
    }
}