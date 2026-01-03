using System.ComponentModel;
using System.Text.Json;
using Aspose.Slides;
using Aspose.Slides.SlideShow;
using AsposeMcpServer.Core.Helpers;
using AsposeMcpServer.Core.Session;
using ModelContextProtocol.Server;

namespace AsposeMcpServer.Tools.PowerPoint;

/// <summary>
///     Unified tool for managing PowerPoint transitions (set, get, delete)
///     Merges: PptSetTransitionTool, PptGetTransitionTool, PptDeleteTransitionTool
/// </summary>
[McpServerToolType]
public class PptTransitionTool
{
    /// <summary>
    ///     Session manager for document session handling.
    /// </summary>
    private readonly DocumentSessionManager? _sessionManager;

    /// <summary>
    ///     Initializes a new instance of the <see cref="PptTransitionTool" /> class.
    /// </summary>
    /// <param name="sessionManager">Optional session manager for in-memory editing.</param>
    public PptTransitionTool(DocumentSessionManager? sessionManager = null)
    {
        _sessionManager = sessionManager;
    }

    [McpServerTool(Name = "ppt_transition")]
    [Description(@"Manage PowerPoint transitions. Supports 3 operations: set, get, delete.

Transition types: Fade, Push, Wipe, Split, Random, Circle, Plus, Diamond, Comb, Cover, Cut, Dissolve, Zoom, and more (all TransitionType enum values supported).

Usage examples:
- Set transition: ppt_transition(operation='set', path='presentation.pptx', slideIndex=0, transitionType='Fade', advanceAfterSeconds=1.5)
- Get transition: ppt_transition(operation='get', path='presentation.pptx', slideIndex=0)
- Delete transition: ppt_transition(operation='delete', path='presentation.pptx', slideIndex=0)")]
    public string Execute(
        [Description(@"Operation to perform.
- 'set': Set slide transition (required params: path, slideIndex, transitionType)
- 'get': Get slide transition (required params: path, slideIndex)
- 'delete': Delete slide transition (required params: path, slideIndex)")]
        string operation,
        [Description("Presentation file path (required if no sessionId)")]
        string? path = null,
        [Description("Session ID for in-memory editing")]
        string? sessionId = null,
        [Description("Output file path (optional, for set/delete operations, defaults to input path)")]
        string? outputPath = null,
        [Description("Slide index (0-based)")] int slideIndex = 0,
        [Description(
            "Transition type: all TransitionType enum values supported (Fade, Push, Wipe, Split, Random, Circle, Plus, Diamond, etc., required for set)")]
        string? transitionType = null,
        [Description("Seconds before auto-advancing to next slide (optional, for set, default: no auto-advance)")]
        double? advanceAfterSeconds = null)
    {
        using var ctx = DocumentContext<Presentation>.Create(_sessionManager, sessionId, path);

        return operation.ToLower() switch
        {
            "set" => SetTransition(ctx, outputPath, slideIndex, transitionType, advanceAfterSeconds),
            "get" => GetTransition(ctx, slideIndex),
            "delete" => DeleteTransition(ctx, outputPath, slideIndex),
            _ => throw new ArgumentException($"Unknown operation: {operation}")
        };
    }

    /// <summary>
    ///     Sets slide transition effect.
    /// </summary>
    /// <param name="ctx">The document context.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="slideIndex">The slide index (0-based).</param>
    /// <param name="transitionTypeStr">The transition type string.</param>
    /// <param name="advanceAfterSeconds">Seconds before auto-advancing to next slide.</param>
    /// <returns>A message indicating the result of the operation.</returns>
    /// <exception cref="ArgumentException">Thrown when transitionType is not provided or invalid.</exception>
    private static string SetTransition(DocumentContext<Presentation> ctx, string? outputPath, int slideIndex,
        string? transitionTypeStr, double? advanceAfterSeconds)
    {
        if (string.IsNullOrEmpty(transitionTypeStr))
            throw new ArgumentException("transitionType is required for set operation");

        if (!Enum.TryParse<TransitionType>(transitionTypeStr, true, out var transitionType))
            throw new ArgumentException(
                $"Invalid transition type: '{transitionTypeStr}'. Valid values: {string.Join(", ", Enum.GetNames<TransitionType>())}");

        var presentation = ctx.Document;
        var slide = PowerPointHelper.GetSlide(presentation, slideIndex);
        var transition = slide.SlideShowTransition;

        transition.Type = transitionType;

        if (advanceAfterSeconds.HasValue)
        {
            transition.AdvanceAfterTime = (uint)(advanceAfterSeconds.Value * 1000);
            transition.AdvanceOnClick = true;
        }

        ctx.Save(outputPath);

        var message = $"Transition '{transition.Type}' set for slide {slideIndex}.";
        if (advanceAfterSeconds.HasValue)
            message += $" Auto-advance after {advanceAfterSeconds.Value:0.##}s.";
        return message + $" {ctx.GetOutputMessage(outputPath)}";
    }

    /// <summary>
    ///     Gets transition information for a slide.
    /// </summary>
    /// <param name="ctx">The document context.</param>
    /// <param name="slideIndex">The slide index (0-based).</param>
    /// <returns>A JSON string containing the transition information.</returns>
    private static string GetTransition(DocumentContext<Presentation> ctx, int slideIndex)
    {
        var presentation = ctx.Document;
        var slide = PowerPointHelper.GetSlide(presentation, slideIndex);
        var transition = slide.SlideShowTransition;

        var advanceAfterTimeMs = transition?.AdvanceAfterTime ?? 0;
        var result = new
        {
            slideIndex,
            hasTransition = transition?.Type != TransitionType.None,
            type = transition?.Type.ToString(),
            speed = transition?.Speed.ToString(),
            advanceOnClick = transition?.AdvanceOnClick,
            advanceAfterSeconds = advanceAfterTimeMs > 0 ? advanceAfterTimeMs / 1000.0 : (double?)null,
            soundMode = transition?.SoundMode.ToString(),
            hasSound = transition?.Sound != null
        };

        return JsonSerializer.Serialize(result, new JsonSerializerOptions { WriteIndented = true });
    }

    /// <summary>
    ///     Removes transition from a slide.
    /// </summary>
    /// <param name="ctx">The document context.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="slideIndex">The slide index (0-based).</param>
    /// <returns>A message indicating the result of the operation.</returns>
    private static string DeleteTransition(DocumentContext<Presentation> ctx, string? outputPath, int slideIndex)
    {
        var presentation = ctx.Document;
        var slide = PowerPointHelper.GetSlide(presentation, slideIndex);
        slide.SlideShowTransition.Type = TransitionType.None;
        slide.SlideShowTransition.AdvanceOnClick = true;
        slide.SlideShowTransition.AdvanceAfterTime = 0;

        ctx.Save(outputPath);

        return $"Transition removed from slide {slideIndex}. {ctx.GetOutputMessage(outputPath)}";
    }
}