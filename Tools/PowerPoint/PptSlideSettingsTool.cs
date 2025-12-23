using System.Text.Json.Nodes;
using Aspose.Slides;
using Aspose.Slides.Export;
using AsposeMcpServer.Core;

namespace AsposeMcpServer.Tools.PowerPoint;

/// <summary>
///     Unified tool for PowerPoint slide settings (set slide size, set slide orientation)
///     Merges: PptSetSlideSizeTool, PptSetSlideOrientationTool
/// </summary>
public class PptSlideSettingsTool : IAsposeTool
{
    public string Description => @"Manage PowerPoint slide settings. Supports 2 operations: set_size, set_orientation.

Usage examples:
- Set slide size: ppt_slide_settings(operation='set_size', path='presentation.pptx', preset='OnScreen16x9')
- Set custom size: ppt_slide_settings(operation='set_size', path='presentation.pptx', preset='Custom', width=960, height=720)
- Set orientation: ppt_slide_settings(operation='set_orientation', path='presentation.pptx', orientation='Portrait')";

    public object InputSchema => new
    {
        type = "object",
        properties = new
        {
            operation = new
            {
                type = "string",
                description = @"Operation to perform.
- 'set_size': Set slide size (required params: path, preset)
- 'set_orientation': Set slide orientation (required params: path, orientation)",
                @enum = new[] { "set_size", "set_orientation" }
            },
            path = new
            {
                type = "string",
                description = "Presentation file path (required for all operations)"
            },
            preset = new
            {
                type = "string",
                description = "Preset: OnScreen16x9, OnScreen16x10, Letter, A4, Banner, Custom (optional, for set_size)"
            },
            width = new
            {
                type = "number",
                description = "Custom width (points) when preset=Custom (optional, for set_size)"
            },
            height = new
            {
                type = "number",
                description = "Custom height (points) when preset=Custom (optional, for set_size)"
            },
            orientation = new
            {
                type = "string",
                description = "Orientation: 'Portrait' or 'Landscape' (required for set_orientation)",
                @enum = new[] { "Portrait", "Landscape" }
            },
            outputPath = new
            {
                type = "string",
                description = "Output file path (optional, for set_orientation operation, defaults to input path)"
            }
        },
        required = new[] { "operation", "path" }
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

        return operation.ToLower() switch
        {
            "set_size" => await SetSlideSizeAsync(arguments, path),
            "set_orientation" => await SetSlideOrientationAsync(arguments, path),
            _ => throw new ArgumentException($"Unknown operation: {operation}")
        };
    }

    /// <summary>
    ///     Sets slide size
    /// </summary>
    /// <param name="arguments">JSON arguments containing width, height, optional outputPath</param>
    /// <param name="path">PowerPoint file path</param>
    /// <returns>Success message</returns>
    private Task<string> SetSlideSizeAsync(JsonObject? arguments, string path)
    {
        return Task.Run(() =>
        {
            var preset = ArgumentHelper.GetString(arguments, "preset", "OnScreen16x9");
            var width = ArgumentHelper.GetDoubleNullable(arguments, "width");
            var height = ArgumentHelper.GetDoubleNullable(arguments, "height");

            using var presentation = new Presentation(path);
            var slideSize = presentation.SlideSize;
            var type = preset.ToLower() switch
            {
                "onscreen16x10" => SlideSizeType.OnScreen16x10,
                "a4" => SlideSizeType.A4Paper,
                "banner" => SlideSizeType.Banner,
                "custom" => SlideSizeType.Custom,
                _ => SlideSizeType.OnScreen
            };

            if (type == SlideSizeType.Custom)
            {
                if (!width.HasValue || !height.HasValue)
                    throw new ArgumentException("custom size requires width and height");
                slideSize.SetSize((float)width.Value, (float)height.Value, SlideSizeScaleType.DoNotScale);
            }
            else
            {
                slideSize.SetSize(type, SlideSizeScaleType.DoNotScale);
            }

            var outputPath = ArgumentHelper.GetAndValidateOutputPath(arguments, path);
            presentation.Save(outputPath, SaveFormat.Pptx);
            return
                $"Slide size set: {slideSize.Type} {(slideSize.Type == SlideSizeType.Custom ? $"{slideSize.Size.Width}x{slideSize.Size.Height}" : string.Empty)}";
        });
    }

    /// <summary>
    ///     Sets slide orientation
    /// </summary>
    /// <param name="arguments">JSON arguments containing orientation (portrait/landscape), optional outputPath</param>
    /// <param name="path">PowerPoint file path</param>
    /// <returns>Success message</returns>
    private Task<string> SetSlideOrientationAsync(JsonObject? arguments, string path)
    {
        return Task.Run(() =>
        {
            var orientation = ArgumentHelper.GetString(arguments, "orientation");

            using var presentation = new Presentation(path);

            presentation.SlideSize.SetSize(
                orientation.ToLower() == "portrait" ? SlideSizeType.A4Paper : SlideSizeType.OnScreen16x10,
                SlideSizeScaleType.EnsureFit);

            var outputPath = ArgumentHelper.GetAndValidateOutputPath(arguments, path);
            presentation.Save(outputPath, SaveFormat.Pptx);

            return $"Slide orientation set to {orientation}: {outputPath}";
        });
    }
}