using System.Text.Json;
using System.Text.Json.Nodes;
using Aspose.Slides;
using Aspose.Slides.Export;
using AsposeMcpServer.Core;

namespace AsposeMcpServer.Tools.PowerPoint;

/// <summary>
///     Unified tool for managing PowerPoint layouts.
///     Supports: set, get_layouts, get_masters, apply_master, apply_layout_range, apply_theme
/// </summary>
public class PptLayoutTool : IAsposeTool
{
    private static readonly Dictionary<string, SlideLayoutType> LayoutTypeMap = new(StringComparer.OrdinalIgnoreCase)
    {
        ["title"] = SlideLayoutType.Title,
        ["titleonly"] = SlideLayoutType.TitleOnly,
        ["blank"] = SlideLayoutType.Blank,
        ["twocolumn"] = SlideLayoutType.TwoColumnText,
        ["twocolumntext"] = SlideLayoutType.TwoColumnText,
        ["sectionheader"] = SlideLayoutType.SectionHeader,
        ["titleandcontent"] = SlideLayoutType.TitleAndObject,
        ["titleandobject"] = SlideLayoutType.TitleAndObject,
        ["objectandtext"] = SlideLayoutType.ObjectAndText,
        ["pictureandcaption"] = SlideLayoutType.PictureAndCaption
    };

    private static readonly string SupportedLayoutTypes = string.Join(", ", LayoutTypeMap.Keys);

    public string Description =>
        @"Manage PowerPoint layouts. Supports 6 operations: set, get_layouts, get_masters, apply_master, apply_layout_range, apply_theme.

Usage examples:
- Set layout: ppt_layout(operation='set', path='presentation.pptx', slideIndex=0, layout='Title')
- Get layouts: ppt_layout(operation='get_layouts', path='presentation.pptx', masterIndex=0)
- Get masters: ppt_layout(operation='get_masters', path='presentation.pptx')
- Apply master: ppt_layout(operation='apply_master', path='presentation.pptx', slideIndices=[0,1,2], masterIndex=0, layoutIndex=0)
- Apply layout range: ppt_layout(operation='apply_layout_range', path='presentation.pptx', slideIndices=[0,1,2], layout='Title')
- Apply theme: ppt_layout(operation='apply_theme', path='presentation.pptx', themePath='theme.potx')";

    public object InputSchema => new
    {
        type = "object",
        properties = new
        {
            operation = new
            {
                type = "string",
                description = @"Operation to perform.
- 'set': Set slide layout (required params: path, slideIndex, layout)
- 'get_layouts': Get available layouts (required params: path)
- 'get_masters': Get master slides (required params: path)
- 'apply_master': Apply master to slides (required params: path, masterIndex, layoutIndex)
- 'apply_layout_range': Apply layout to multiple slides (required params: path, slideIndices, layout)
- 'apply_theme': Apply theme template by copying master slides (required params: path, themePath)",
                @enum = new[]
                    { "set", "get_layouts", "get_masters", "apply_master", "apply_layout_range", "apply_theme" }
            },
            path = new
            {
                type = "string",
                description = "Presentation file path (required for all operations)"
            },
            slideIndex = new
            {
                type = "number",
                description = "Slide index (0-based, required for set)"
            },
            layout = new
            {
                type = "string",
                description =
                    "Layout type: Title, TitleOnly, Blank, TwoColumn, SectionHeader, TitleAndContent, ObjectAndText, PictureAndCaption (required for set/apply_layout_range)"
            },
            masterIndex = new
            {
                type = "number",
                description = "Master index (0-based, optional for get_layouts, required for apply_master)"
            },
            layoutIndex = new
            {
                type = "number",
                description = "Layout index under master (0-based, required for apply_master)"
            },
            slideIndices = new
            {
                type = "array",
                items = new { type = "number" },
                description = "Slide indices array (required for apply_layout_range, optional for apply_master)"
            },
            themePath = new
            {
                type = "string",
                description = "Theme template file path (.potx/.pptx, required for apply_theme)"
            },
            outputPath = new
            {
                type = "string",
                description = "Output file path (optional, defaults to input path)"
            }
        },
        required = new[] { "operation", "path" }
    };

    /// <summary>
    ///     Executes the tool operation with the provided JSON arguments.
    /// </summary>
    /// <param name="arguments">JSON arguments object containing operation parameters.</param>
    /// <returns>Result message as a string.</returns>
    /// <exception cref="ArgumentException">Thrown when operation is unknown.</exception>
    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var operation = ArgumentHelper.GetString(arguments, "operation");
        var path = ArgumentHelper.GetAndValidatePath(arguments);
        var outputPath = ArgumentHelper.GetAndValidateOutputPath(arguments, path);

        return operation.ToLower() switch
        {
            "set" => await SetLayoutAsync(path, outputPath, arguments),
            "get_layouts" => await GetLayoutsAsync(path, arguments),
            "get_masters" => await GetMastersAsync(path),
            "apply_master" => await ApplyMasterAsync(path, outputPath, arguments),
            "apply_layout_range" => await ApplyLayoutRangeAsync(path, outputPath, arguments),
            "apply_theme" => await ApplyThemeAsync(path, outputPath, arguments),
            _ => throw new ArgumentException($"Unknown operation: {operation}")
        };
    }

    /// <summary>
    ///     Sets layout for a slide.
    /// </summary>
    /// <param name="path">PowerPoint file path.</param>
    /// <param name="outputPath">Output file path.</param>
    /// <param name="arguments">JSON arguments containing slideIndex, layout.</param>
    /// <returns>Success message.</returns>
    /// <exception cref="ArgumentException">Thrown when slideIndex is out of range or layout type is not supported.</exception>
    private Task<string> SetLayoutAsync(string path, string outputPath, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var slideIndex = ArgumentHelper.GetInt(arguments, "slideIndex");
            var layoutStr = ArgumentHelper.GetString(arguments, "layout");

            using var presentation = new Presentation(path);
            PowerPointHelper.ValidateCollectionIndex(slideIndex, presentation.Slides.Count, "slide");

            var layout = FindLayoutByType(presentation, layoutStr);
            presentation.Slides[slideIndex].LayoutSlide = layout;

            presentation.Save(outputPath, SaveFormat.Pptx);
            return $"Layout '{layoutStr}' set for slide {slideIndex}. Output: {outputPath}";
        });
    }

    /// <summary>
    ///     Gets available layouts with layout type information.
    /// </summary>
    /// <param name="path">PowerPoint file path.</param>
    /// <param name="arguments">JSON arguments containing optional masterIndex.</param>
    /// <returns>JSON string with available layouts including layoutType.</returns>
    /// <exception cref="ArgumentException">Thrown when masterIndex is out of range.</exception>
    private Task<string> GetLayoutsAsync(string path, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var masterIndex = ArgumentHelper.GetIntNullable(arguments, "masterIndex");

            using var presentation = new Presentation(path);

            if (masterIndex.HasValue)
            {
                PowerPointHelper.ValidateCollectionIndex(masterIndex.Value, presentation.Masters.Count, "master");

                var master = presentation.Masters[masterIndex.Value];
                var layoutsList = BuildLayoutsList(master.LayoutSlides);

                var result = new
                {
                    masterIndex = masterIndex.Value,
                    count = master.LayoutSlides.Count,
                    layouts = layoutsList
                };

                return JsonSerializer.Serialize(result, new JsonSerializerOptions { WriteIndented = true });
            }
            else
            {
                var mastersList = new List<object>();

                for (var i = 0; i < presentation.Masters.Count; i++)
                {
                    var master = presentation.Masters[i];
                    mastersList.Add(new
                    {
                        masterIndex = i,
                        layoutCount = master.LayoutSlides.Count,
                        layouts = BuildLayoutsList(master.LayoutSlides)
                    });
                }

                var result = new
                {
                    mastersCount = presentation.Masters.Count,
                    masters = mastersList
                };

                return JsonSerializer.Serialize(result, new JsonSerializerOptions { WriteIndented = true });
            }
        });
    }

    /// <summary>
    ///     Gets master slides information.
    /// </summary>
    /// <param name="path">PowerPoint file path.</param>
    /// <returns>JSON string with master slides.</returns>
    private Task<string> GetMastersAsync(string path)
    {
        return Task.Run(() =>
        {
            using var presentation = new Presentation(path);

            if (presentation.Masters.Count == 0)
            {
                var emptyResult = new
                {
                    count = 0,
                    masters = Array.Empty<object>(),
                    message = "No master slides found"
                };
                return JsonSerializer.Serialize(emptyResult, new JsonSerializerOptions { WriteIndented = true });
            }

            var mastersList = new List<object>();

            for (var i = 0; i < presentation.Masters.Count; i++)
            {
                var master = presentation.Masters[i];
                mastersList.Add(new
                {
                    index = i,
                    name = master.Name,
                    layoutCount = master.LayoutSlides.Count,
                    layouts = BuildLayoutsList(master.LayoutSlides)
                });
            }

            var result = new
            {
                count = presentation.Masters.Count,
                masters = mastersList
            };

            return JsonSerializer.Serialize(result, new JsonSerializerOptions { WriteIndented = true });
        });
    }

    /// <summary>
    ///     Applies a master slide layout to specified slides.
    /// </summary>
    /// <param name="path">PowerPoint file path.</param>
    /// <param name="outputPath">Output file path.</param>
    /// <param name="arguments">JSON arguments containing masterIndex, layoutIndex, optional slideIndices.</param>
    /// <returns>Success message.</returns>
    /// <exception cref="ArgumentException">Thrown when masterIndex, layoutIndex, or slideIndices is out of range.</exception>
    private Task<string> ApplyMasterAsync(string path, string outputPath, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var slideIndicesArray = ArgumentHelper.GetArray(arguments, "slideIndices", false);
            var slideIndices = slideIndicesArray?.Select(x => x?.GetValue<int>() ?? -1).ToArray();
            var masterIndex = ArgumentHelper.GetInt(arguments, "masterIndex", 0);
            var layoutIndex = ArgumentHelper.GetInt(arguments, "layoutIndex", 0);

            using var presentation = new Presentation(path);

            PowerPointHelper.ValidateCollectionIndex(masterIndex, presentation.Masters.Count, "master");
            var master = presentation.Masters[masterIndex];
            PowerPointHelper.ValidateCollectionIndex(layoutIndex, master.LayoutSlides.Count, "layout");

            var targets = slideIndices?.Length > 0
                ? slideIndices
                : Enumerable.Range(0, presentation.Slides.Count).ToArray();

            ValidateSlideIndices(targets, presentation.Slides.Count);

            var layout = master.LayoutSlides[layoutIndex];
            foreach (var idx in targets)
                presentation.Slides[idx].LayoutSlide = layout;

            presentation.Save(outputPath, SaveFormat.Pptx);
            return
                $"Master {masterIndex} / Layout {layoutIndex} applied to {targets.Length} slides. Output: {outputPath}";
        });
    }

    /// <summary>
    ///     Applies layout to a range of slides.
    /// </summary>
    /// <param name="path">PowerPoint file path.</param>
    /// <param name="outputPath">Output file path.</param>
    /// <param name="arguments">JSON arguments containing layout, slideIndices.</param>
    /// <returns>Success message.</returns>
    /// <exception cref="ArgumentException">Thrown when slideIndices is out of range or layout type is not supported.</exception>
    private Task<string> ApplyLayoutRangeAsync(string path, string outputPath, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var layoutStr = ArgumentHelper.GetString(arguments, "layout");
            var slideIndicesArray = ArgumentHelper.GetArray(arguments, "slideIndices");
            var slideIndices = slideIndicesArray.Select(x => x?.GetValue<int>() ?? -1).ToArray();

            using var presentation = new Presentation(path);

            ValidateSlideIndices(slideIndices, presentation.Slides.Count);

            var layout = FindLayoutByType(presentation, layoutStr);

            foreach (var idx in slideIndices)
                presentation.Slides[idx].LayoutSlide = layout;

            presentation.Save(outputPath, SaveFormat.Pptx);
            return $"Layout '{layoutStr}' applied to {slideIndices.Length} slide(s). Output: {outputPath}";
        });
    }

    /// <summary>
    ///     Applies a theme to the presentation by copying master slides.
    /// </summary>
    /// <param name="path">PowerPoint file path.</param>
    /// <param name="outputPath">Output file path.</param>
    /// <param name="arguments">JSON arguments containing themePath.</param>
    /// <returns>Success message with copied master count.</returns>
    /// <exception cref="FileNotFoundException">Thrown when theme file is not found.</exception>
    private Task<string> ApplyThemeAsync(string path, string outputPath, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var themePath = ArgumentHelper.GetString(arguments, "themePath");

            if (!File.Exists(themePath))
                throw new FileNotFoundException($"Theme file not found: {themePath}");

            using var presentation = new Presentation(path);
            using var themePresentation = new Presentation(themePath);

            if (themePresentation.Masters.Count == 0)
                throw new InvalidOperationException("Theme presentation does not contain any master slides.");

            var copiedCount = 0;
            foreach (var themeMaster in themePresentation.Masters)
            {
                presentation.Masters.AddClone(themeMaster);
                copiedCount++;
            }

            if (presentation.Slides.Count > 0 && themePresentation.Masters.Count > 0)
            {
                var newMaster = presentation.Masters[^1];
                if (newMaster.LayoutSlides.Count > 0)
                {
                    var defaultLayout = newMaster.LayoutSlides[0];
                    foreach (var slide in presentation.Slides)
                        slide.LayoutSlide = defaultLayout;
                }
            }

            presentation.Save(outputPath, SaveFormat.Pptx);

            return
                $"Theme applied ({copiedCount} master(s) copied, layout applied to all slides). Output: {outputPath}";
        });
    }

    #region Helper Methods

    /// <summary>
    ///     Finds a layout slide by layout type string.
    /// </summary>
    /// <param name="presentation">The presentation to search in.</param>
    /// <param name="layoutStr">The layout type string.</param>
    /// <returns>The matching layout slide.</returns>
    /// <exception cref="ArgumentException">Thrown when layout type is not found.</exception>
    private static ILayoutSlide FindLayoutByType(IPresentation presentation, string layoutStr)
    {
        if (!LayoutTypeMap.TryGetValue(layoutStr, out var layoutType))
            throw new ArgumentException(
                $"Unknown layout type: '{layoutStr}'. Supported types: {SupportedLayoutTypes}");

        var layout = presentation.LayoutSlides.FirstOrDefault(ls => ls.LayoutType == layoutType);
        if (layout == null)
            throw new ArgumentException(
                $"Layout type '{layoutStr}' not found in this presentation. Use get_layouts to see available layouts.");

        return layout;
    }

    /// <summary>
    ///     Validates all slide indices before processing.
    /// </summary>
    /// <param name="indices">Array of slide indices to validate.</param>
    /// <param name="slideCount">Total number of slides in the presentation.</param>
    /// <exception cref="ArgumentException">Thrown when any slide index is out of range.</exception>
    private static void ValidateSlideIndices(int[] indices, int slideCount)
    {
        var invalidIndices = indices.Where(idx => idx < 0 || idx >= slideCount).ToList();
        if (invalidIndices.Count > 0)
            throw new ArgumentException(
                $"Invalid slide indices: [{string.Join(", ", invalidIndices)}]. Valid range: 0 to {slideCount - 1}");
    }

    /// <summary>
    ///     Builds a list of layout information including layout type.
    /// </summary>
    /// <param name="layoutSlides">The layout slides collection.</param>
    /// <returns>List of layout information objects.</returns>
    private static List<object> BuildLayoutsList(IMasterLayoutSlideCollection layoutSlides)
    {
        var layoutsList = new List<object>();
        for (var i = 0; i < layoutSlides.Count; i++)
        {
            var layout = layoutSlides[i];
            layoutsList.Add(new
            {
                index = i,
                name = layout.Name,
                layoutType = layout.LayoutType.ToString()
            });
        }

        return layoutsList;
    }

    #endregion
}