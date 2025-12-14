using System.Text.Json.Nodes;
using System.Text;
using Aspose.Slides;
using Aspose.Slides.Export;
using Aspose.Slides.SlideShow;
using AsposeMcpServer.Core;

namespace AsposeMcpServer.Tools;

/// <summary>
/// Unified tool for managing PowerPoint transitions (set, get, delete)
/// Merges: PptSetTransitionTool, PptGetTransitionTool, PptDeleteTransitionTool
/// </summary>
public class PptTransitionTool : IAsposeTool
{
    public string Description => @"Manage PowerPoint transitions. Supports 3 operations: set, get, delete.

Usage examples:
- Set transition: ppt_transition(operation='set', path='presentation.pptx', slideIndex=0, transitionType='Fade', durationSeconds=1.5)
- Get transition: ppt_transition(operation='get', path='presentation.pptx', slideIndex=0)
- Delete transition: ppt_transition(operation='delete', path='presentation.pptx', slideIndex=0)";

    public object InputSchema => new
    {
        type = "object",
        properties = new
        {
            operation = new
            {
                type = "string",
                description = @"Operation to perform.
- 'set': Set slide transition (required params: path, slideIndex, transitionType)
- 'get': Get slide transition (required params: path, slideIndex)
- 'delete': Delete slide transition (required params: path, slideIndex)",
                @enum = new[] { "set", "get", "delete" }
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
            transitionType = new
            {
                type = "string",
                description = "Transition type (Fade, Push, Wipe, Split, RandomBars, etc., required for set)"
            },
            durationSeconds = new
            {
                type = "number",
                description = "Transition duration in seconds (optional, for set, default 1.0)"
            }
        },
        required = new[] { "operation", "path", "slideIndex" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var operation = ArgumentHelper.GetString(arguments, "operation", "operation");
        var path = ArgumentHelper.GetAndValidatePath(arguments);
        var slideIndex = ArgumentHelper.GetInt(arguments, "slideIndex", "slideIndex");

        return operation.ToLower() switch
        {
            "set" => await SetTransitionAsync(arguments, path, slideIndex),
            "get" => await GetTransitionAsync(arguments, path, slideIndex),
            "delete" => await DeleteTransitionAsync(arguments, path, slideIndex),
            _ => throw new ArgumentException($"Unknown operation: {operation}")
        };
    }

    /// <summary>
    /// Sets slide transition effect
    /// </summary>
    /// <param name="arguments">JSON arguments containing transitionType, optional duration, outputPath</param>
    /// <param name="path">PowerPoint file path</param>
    /// <param name="slideIndex">Slide index (0-based)</param>
    /// <returns>Success message</returns>
    private async Task<string> SetTransitionAsync(JsonObject? arguments, string path, int slideIndex)
    {
        var transitionTypeStr = ArgumentHelper.GetString(arguments, "transitionType", "transitionType");
        var duration = arguments?["durationSeconds"]?.GetValue<double?>() ?? 1.0;

        using var presentation = new Presentation(path);
        var slide = PowerPointHelper.GetSlide(presentation, slideIndex);
        var transition = slide.SlideShowTransition;
        transition.Type = transitionTypeStr.ToLower() switch
        {
            "push" => TransitionType.Push,
            "wipe" => TransitionType.Wipe,
            "split" => TransitionType.Split,
            "randombars" => TransitionType.Random,
            "circle" => TransitionType.Circle,
            "plus" => TransitionType.Plus,
            "diamond" => TransitionType.Diamond,
            "fade" => TransitionType.Fade,
            _ => TransitionType.Fade
        };
        transition.AdvanceAfterTime = (uint)(duration * 1000);

        presentation.Save(path, SaveFormat.Pptx);
        return await Task.FromResult($"已設定投影片 {slideIndex} 轉場：{transition.Type}，時間 {duration:0.##}s");
    }

    /// <summary>
    /// Gets transition information for a slide
    /// </summary>
    /// <param name="arguments">JSON arguments (no specific parameters required)</param>
    /// <param name="path">PowerPoint file path</param>
    /// <param name="slideIndex">Slide index (0-based)</param>
    /// <returns>Formatted string with transition details</returns>
    private async Task<string> GetTransitionAsync(JsonObject? arguments, string path, int slideIndex)
    {
        using var presentation = new Presentation(path);
        var slide = PowerPointHelper.GetSlide(presentation, slideIndex);
        var transition = slide.SlideShowTransition;
        var sb = new StringBuilder();

        sb.AppendLine($"=== Slide {slideIndex} Transition ===");
        if (transition != null)
        {
            sb.AppendLine($"Type: {transition.Type}");
            sb.AppendLine($"Speed: {transition.Speed}");
            sb.AppendLine($"AdvanceOnClick: {transition.AdvanceOnClick}");
            sb.AppendLine($"AdvanceAfterTime: {transition.AdvanceAfterTime}ms");
            sb.AppendLine($"SoundMode: {transition.SoundMode}");
            if (transition.Sound != null)
            {
                sb.AppendLine($"Sound: {transition.Sound}");
            }
        }
        else
        {
            sb.AppendLine("No transition set");
        }

        return await Task.FromResult(sb.ToString());
    }

    /// <summary>
    /// Removes transition from a slide
    /// </summary>
    /// <param name="arguments">JSON arguments containing optional outputPath</param>
    /// <param name="path">PowerPoint file path</param>
    /// <param name="slideIndex">Slide index (0-based)</param>
    /// <returns>Success message</returns>
    private async Task<string> DeleteTransitionAsync(JsonObject? arguments, string path, int slideIndex)
    {
        using var presentation = new Presentation(path);
        var slide = PowerPointHelper.GetSlide(presentation, slideIndex);
        slide.SlideShowTransition.Type = TransitionType.Fade;
        slide.SlideShowTransition.AdvanceOnClick = false;
        slide.SlideShowTransition.AdvanceAfterTime = 0;

        presentation.Save(path, SaveFormat.Pptx);

        return await Task.FromResult($"Transition removed from slide {slideIndex}: {path}");
    }
}

