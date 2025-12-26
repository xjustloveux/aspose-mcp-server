using System.Text.Json;
using System.Text.Json.Nodes;
using Aspose.Slides;
using Aspose.Slides.Export;
using Aspose.Slides.SlideShow;
using AsposeMcpServer.Core;

namespace AsposeMcpServer.Tools.PowerPoint;

/// <summary>
///     Unified tool for managing PowerPoint transitions (set, get, delete)
///     Merges: PptSetTransitionTool, PptGetTransitionTool, PptDeleteTransitionTool
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
            },
            outputPath = new
            {
                type = "string",
                description = "Output file path (optional, for set/remove operations, defaults to input path)"
            }
        },
        required = new[] { "operation", "path", "slideIndex" }
    };

    /// <summary>
    ///     Executes the tool operation with the provided JSON arguments
    /// </summary>
    /// <param name="arguments">JSON arguments object containing operation parameters</param>
    /// <returns>Result message as a string</returns>
    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var operation = ArgumentHelper.GetString(arguments, "operation");
        var path = ArgumentHelper.GetAndValidatePath(arguments);
        var outputPath = ArgumentHelper.GetAndValidateOutputPath(arguments, path);
        var slideIndex = ArgumentHelper.GetInt(arguments, "slideIndex");

        return operation.ToLower() switch
        {
            "set" => await SetTransitionAsync(path, outputPath, slideIndex, arguments),
            "get" => await GetTransitionAsync(path, slideIndex),
            "delete" => await DeleteTransitionAsync(path, outputPath, slideIndex),
            _ => throw new ArgumentException($"Unknown operation: {operation}")
        };
    }

    /// <summary>
    ///     Sets slide transition effect
    /// </summary>
    /// <param name="path">PowerPoint file path</param>
    /// <param name="outputPath">Output file path</param>
    /// <param name="slideIndex">Slide index (0-based)</param>
    /// <param name="arguments">JSON arguments containing transitionType, optional duration</param>
    /// <returns>Success message</returns>
    private Task<string> SetTransitionAsync(string path, string outputPath, int slideIndex, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var transitionTypeStr = ArgumentHelper.GetString(arguments, "transitionType");
            var duration = ArgumentHelper.GetDouble(arguments, "durationSeconds", 1.0);

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
                _ => TransitionType.Fade
            };
            transition.AdvanceAfterTime = (uint)(duration * 1000);

            presentation.Save(outputPath, SaveFormat.Pptx);
            return
                $"Transition '{transition.Type}' set for slide {slideIndex} (duration {duration:0.##}s). Output: {outputPath}";
        });
    }

    /// <summary>
    ///     Gets transition information for a slide
    /// </summary>
    /// <param name="path">PowerPoint file path</param>
    /// <param name="slideIndex">Slide index (0-based)</param>
    /// <returns>JSON string with transition details</returns>
    private Task<string> GetTransitionAsync(string path, int slideIndex)
    {
        return Task.Run(() =>
        {
            using var presentation = new Presentation(path);
            var slide = PowerPointHelper.GetSlide(presentation, slideIndex);
            var transition = slide.SlideShowTransition;

            var result = new
            {
                slideIndex,
                hasTransition = transition != null,
                type = transition?.Type.ToString(),
                speed = transition?.Speed.ToString(),
                advanceOnClick = transition?.AdvanceOnClick,
                advanceAfterTimeMs = transition?.AdvanceAfterTime,
                soundMode = transition?.SoundMode.ToString(),
                hasSound = transition?.Sound != null
            };

            return JsonSerializer.Serialize(result, new JsonSerializerOptions { WriteIndented = true });
        });
    }

    /// <summary>
    ///     Removes transition from a slide
    /// </summary>
    /// <param name="path">PowerPoint file path</param>
    /// <param name="outputPath">Output file path</param>
    /// <param name="slideIndex">Slide index (0-based)</param>
    /// <returns>Success message</returns>
    private Task<string> DeleteTransitionAsync(string path, string outputPath, int slideIndex)
    {
        return Task.Run(() =>
        {
            using var presentation = new Presentation(path);
            var slide = PowerPointHelper.GetSlide(presentation, slideIndex);
            slide.SlideShowTransition.Type = TransitionType.Fade;
            slide.SlideShowTransition.AdvanceOnClick = false;
            slide.SlideShowTransition.AdvanceAfterTime = 0;

            presentation.Save(outputPath, SaveFormat.Pptx);

            return $"Transition removed from slide {slideIndex}. Output: {outputPath}";
        });
    }
}