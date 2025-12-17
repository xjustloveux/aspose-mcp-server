using System.Text;
using System.Text.Json.Nodes;
using Aspose.Slides;
using Aspose.Slides.Export;
using AsposeMcpServer.Core;

namespace AsposeMcpServer.Tools.PowerPoint;

/// <summary>
///     Unified tool for managing PowerPoint layouts (set, get layouts, get masters, apply master, apply layout range,
///     apply theme)
///     Merges: PptSetLayoutTool, PptGetLayoutsTool, PptGetMasterSlidesTool, PptApplyMasterTool,
///     PptApplyLayoutRangeTool, PptApplyThemeTool
/// </summary>
public class PptLayoutTool : IAsposeTool
{
    public string Description =>
        @"Manage PowerPoint layouts. Supports 6 operations: set, get_layouts, get_masters, apply_master, apply_layout_range, apply_theme.

Usage examples:
- Set layout: ppt_layout(operation='set', path='presentation.pptx', slideIndex=0, layout='Title')
- Get layouts: ppt_layout(operation='get_layouts', path='presentation.pptx', masterIndex=0)
- Get masters: ppt_layout(operation='get_masters', path='presentation.pptx')
- Apply master: ppt_layout(operation='apply_master', path='presentation.pptx', slideIndex=0, masterIndex=0, layoutIndex=0)
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
- 'apply_master': Apply master to slide (required params: path, slideIndex, masterIndex, layoutIndex)
- 'apply_layout_range': Apply layout to multiple slides (required params: path, slideIndices, layout)
- 'apply_theme': Apply theme template (required params: path, themePath)",
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
                    "Layout type (Title, TitleOnly, Blank, TwoColumn, SectionHeader, etc., required for set/apply_layout_range)"
            },
            masterIndex = new
            {
                type = "number",
                description = "Master index (0-based, optional, for get_layouts/apply_master)"
            },
            layoutIndex = new
            {
                type = "number",
                description = "Layout index under master (0-based, optional, for apply_master)"
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
                description = "Theme template file path (required for apply_theme)"
            },
            outputPath = new
            {
                type = "string",
                description = "Output file path (optional, for apply_theme operation, defaults to input path)"
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
            "set" => await SetLayoutAsync(arguments, path),
            "get_layouts" => await GetLayoutsAsync(arguments, path),
            "get_masters" => await GetMastersAsync(arguments, path),
            "apply_master" => await ApplyMasterAsync(arguments, path),
            "apply_layout_range" => await ApplyLayoutRangeAsync(arguments, path),
            "apply_theme" => await ApplyThemeAsync(arguments, path),
            _ => throw new ArgumentException($"Unknown operation: {operation}")
        };
    }

    /// <summary>
    ///     Sets layout for a slide
    /// </summary>
    /// <param name="arguments">JSON arguments containing slideIndex, layoutType, optional outputPath</param>
    /// <param name="path">PowerPoint file path</param>
    /// <returns>Success message</returns>
    private async Task<string> SetLayoutAsync(JsonObject? arguments, string path)
    {
        var slideIndex = ArgumentHelper.GetInt(arguments, "slideIndex");
        var layoutStr = ArgumentHelper.GetString(arguments, "layout");

        using var presentation = new Presentation(path);
        if (slideIndex < 0 || slideIndex >= presentation.Slides.Count)
            throw new ArgumentException($"slideIndex must be between 0 and {presentation.Slides.Count - 1}");

        var layoutType = layoutStr.ToLower() switch
        {
            "title" => SlideLayoutType.Title,
            "titleonly" => SlideLayoutType.TitleOnly,
            "blank" => SlideLayoutType.Blank,
            "twocolumn" => SlideLayoutType.TwoColumnText,
            "sectionheader" => SlideLayoutType.SectionHeader,
            _ => SlideLayoutType.Custom
        };

        var layout = presentation.LayoutSlides.FirstOrDefault(ls => ls.LayoutType == layoutType) ??
                     presentation.LayoutSlides[0];
        presentation.Slides[slideIndex].LayoutSlide = layout;

        var outputPath = ArgumentHelper.GetAndValidateOutputPath(arguments, path);
        presentation.Save(outputPath, SaveFormat.Pptx);
        return await Task.FromResult($"Layout set for slide {slideIndex}: {layoutStr}");
    }

    /// <summary>
    ///     Gets available layouts
    /// </summary>
    /// <param name="arguments">JSON arguments (no specific parameters required)</param>
    /// <param name="path">PowerPoint file path</param>
    /// <returns>Formatted string with available layouts</returns>
    private async Task<string> GetLayoutsAsync(JsonObject? arguments, string path)
    {
        var masterIndex = ArgumentHelper.GetIntNullable(arguments, "masterIndex");

        using var presentation = new Presentation(path);
        var sb = new StringBuilder();

        if (masterIndex.HasValue)
        {
            if (masterIndex.Value < 0 || masterIndex.Value >= presentation.Masters.Count)
                throw new ArgumentException($"masterIndex must be between 0 and {presentation.Masters.Count - 1}");
            var master = presentation.Masters[masterIndex.Value];
            sb.AppendLine($"=== Master {masterIndex.Value} Layouts ===");
            sb.AppendLine($"Total: {master.LayoutSlides.Count}");
            for (var i = 0; i < master.LayoutSlides.Count; i++)
            {
                var layout = master.LayoutSlides[i];
                sb.AppendLine($"  [{i}] {layout.Name ?? "(unnamed)"}");
            }
        }
        else
        {
            sb.AppendLine("=== All Layouts ===");
            for (var i = 0; i < presentation.Masters.Count; i++)
            {
                var master = presentation.Masters[i];
                sb.AppendLine($"\nMaster {i}: {master.LayoutSlides.Count} layout(s)");
                for (var j = 0; j < master.LayoutSlides.Count; j++)
                {
                    var layout = master.LayoutSlides[j];
                    sb.AppendLine($"  [{j}] {layout.Name ?? "(unnamed)"}");
                }
            }
        }

        return await Task.FromResult(sb.ToString());
    }

    /// <summary>
    ///     Gets master slides information
    /// </summary>
    /// <param name="_">Unused parameter</param>
    /// <param name="path">PowerPoint file path</param>
    /// <returns>Formatted string with master slides</returns>
    private async Task<string> GetMastersAsync(JsonObject? _, string path)
    {
        using var presentation = new Presentation(path);
        var sb = new StringBuilder();

        sb.AppendLine("=== Master Slides ===");
        sb.AppendLine($"Total: {presentation.Masters.Count}");

        for (var i = 0; i < presentation.Masters.Count; i++)
        {
            var master = presentation.Masters[i];
            sb.AppendLine($"\nMaster {i}:");
            sb.AppendLine($"  Name: {master.Name ?? "(unnamed)"}");
            sb.AppendLine($"  Layouts: {master.LayoutSlides.Count}");
            for (var j = 0; j < master.LayoutSlides.Count; j++)
            {
                var layout = master.LayoutSlides[j];
                sb.AppendLine($"    [{j}] {layout.Name ?? "(unnamed)"}");
            }
        }

        return await Task.FromResult(sb.ToString());
    }

    /// <summary>
    ///     Applies a master slide to slides
    /// </summary>
    /// <param name="arguments">JSON arguments containing masterIndex, optional slideIndexes, outputPath</param>
    /// <param name="path">PowerPoint file path</param>
    /// <returns>Success message</returns>
    private async Task<string> ApplyMasterAsync(JsonObject? arguments, string path)
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

        foreach (var idx in targets)
            if (idx < 0 || idx >= presentation.Slides.Count)
                throw new ArgumentException($"slide index {idx} out of range");

        var layout = master.LayoutSlides[layoutIndex];
        foreach (var idx in targets) presentation.Slides[idx].LayoutSlide = layout;

        var outputPath = ArgumentHelper.GetAndValidateOutputPath(arguments, path);
        presentation.Save(outputPath, SaveFormat.Pptx);
        return await Task.FromResult($"Master {masterIndex} / Layout {layoutIndex} applied to {targets.Length} slides");
    }

    /// <summary>
    ///     Applies layout to a range of slides
    /// </summary>
    /// <param name="arguments">JSON arguments containing layoutType, startIndex, endIndex, optional outputPath</param>
    /// <param name="path">PowerPoint file path</param>
    /// <returns>Success message</returns>
    private async Task<string> ApplyLayoutRangeAsync(JsonObject? arguments, string path)
    {
        var layoutStr = ArgumentHelper.GetString(arguments, "layout");
        var slideIndicesArray = ArgumentHelper.GetArray(arguments, "slideIndices");

        var slideIndices = slideIndicesArray.Select(x => x?.GetValue<int>() ?? -1).ToArray();

        using var presentation = new Presentation(path);

        foreach (var idx in slideIndices)
            if (idx < 0 || idx >= presentation.Slides.Count)
                throw new ArgumentException($"slide index {idx} out of range");

        var layoutType = layoutStr.ToLower() switch
        {
            "title" => SlideLayoutType.Title,
            "titleonly" => SlideLayoutType.TitleOnly,
            "blank" => SlideLayoutType.Blank,
            "twocolumn" => SlideLayoutType.TwoColumnText,
            "sectionheader" => SlideLayoutType.SectionHeader,
            _ => SlideLayoutType.Custom
        };

        var layout = presentation.LayoutSlides.FirstOrDefault(ls => ls.LayoutType == layoutType) ??
                     presentation.LayoutSlides[0];

        foreach (var idx in slideIndices) presentation.Slides[idx].LayoutSlide = layout;

        var outputPath = ArgumentHelper.GetAndValidateOutputPath(arguments, path);
        presentation.Save(outputPath, SaveFormat.Pptx);
        return await Task.FromResult($"Layout {layoutStr} applied to {slideIndices.Length} slides");
    }

    /// <summary>
    ///     Applies a theme to the presentation
    /// </summary>
    /// <param name="arguments">JSON arguments containing themePath, optional outputPath</param>
    /// <param name="path">PowerPoint file path</param>
    /// <returns>Success message</returns>
    private async Task<string> ApplyThemeAsync(JsonObject? arguments, string path)
    {
        var themePath = ArgumentHelper.GetString(arguments, "themePath");

        using var presentation = new Presentation(path);
        using var themePresentation = new Presentation(themePath);

        // Copy theme from the first slide of theme presentation
        presentation.Slides[0].LayoutSlide = themePresentation.Slides[0].LayoutSlide;

        var outputPath = ArgumentHelper.GetAndValidateOutputPath(arguments, path);
        presentation.Save(outputPath, SaveFormat.Pptx);

        return await Task.FromResult($"Theme applied to presentation: {outputPath}");
    }
}