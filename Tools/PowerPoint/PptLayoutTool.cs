using System.ComponentModel;
using System.Text.Json;
using Aspose.Slides;
using AsposeMcpServer.Core.Helpers;
using AsposeMcpServer.Core.Session;
using ModelContextProtocol.Server;

namespace AsposeMcpServer.Tools.PowerPoint;

/// <summary>
///     Unified tool for managing PowerPoint layouts.
///     Supports: set, get_layouts, get_masters, apply_master, apply_layout_range, apply_theme
/// </summary>
[McpServerToolType]
public class PptLayoutTool
{
    /// <summary>
    ///     Mapping of layout type string names to SlideLayoutType enum values.
    /// </summary>
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

    /// <summary>
    ///     Comma-separated list of supported layout type names for error messages.
    /// </summary>
    private static readonly string SupportedLayoutTypes = string.Join(", ", LayoutTypeMap.Keys);

    /// <summary>
    ///     Session manager for document lifecycle management.
    /// </summary>
    private readonly DocumentSessionManager? _sessionManager;

    /// <summary>
    ///     Initializes a new instance of the <see cref="PptLayoutTool" /> class.
    /// </summary>
    /// <param name="sessionManager">Optional session manager for in-memory document editing.</param>
    public PptLayoutTool(DocumentSessionManager? sessionManager = null)
    {
        _sessionManager = sessionManager;
    }

    [McpServerTool(Name = "ppt_layout")]
    [Description(
        @"Manage PowerPoint layouts. Supports 6 operations: set, get_layouts, get_masters, apply_master, apply_layout_range, apply_theme.

Usage examples:
- Set layout: ppt_layout(operation='set', path='presentation.pptx', slideIndex=0, layout='Title')
- Get layouts: ppt_layout(operation='get_layouts', path='presentation.pptx', masterIndex=0)
- Get masters: ppt_layout(operation='get_masters', path='presentation.pptx')
- Apply master: ppt_layout(operation='apply_master', path='presentation.pptx', slideIndices=[0,1,2], masterIndex=0, layoutIndex=0)
- Apply layout range: ppt_layout(operation='apply_layout_range', path='presentation.pptx', slideIndices=[0,1,2], layout='Title')
- Apply theme: ppt_layout(operation='apply_theme', path='presentation.pptx', themePath='theme.potx')")]
    public string Execute(
        [Description("Operation: set, get_layouts, get_masters, apply_master, apply_layout_range, apply_theme")]
        string operation,
        [Description("Presentation file path (required if no sessionId)")]
        string? path = null,
        [Description("Session ID for in-memory editing")]
        string? sessionId = null,
        [Description("Output file path (file mode only)")]
        string? outputPath = null,
        [Description("Slide index (0-based, required for set)")]
        int? slideIndex = null,
        [Description(
            "Layout type: Title, TitleOnly, Blank, TwoColumn, SectionHeader, TitleAndContent, ObjectAndText, PictureAndCaption")]
        string? layout = null,
        [Description("Master index (0-based, optional for get_layouts, required for apply_master)")]
        int? masterIndex = null,
        [Description("Layout index under master (0-based, required for apply_master)")]
        int? layoutIndex = null,
        [Description("Slide indices array as JSON (required for apply_layout_range, optional for apply_master)")]
        string? slideIndices = null,
        [Description("Theme template file path (.potx/.pptx, required for apply_theme)")]
        string? themePath = null)
    {
        using var ctx = DocumentContext<Presentation>.Create(_sessionManager, sessionId, path);

        return operation.ToLower() switch
        {
            "set" => SetLayout(ctx, outputPath, slideIndex, layout),
            "get_layouts" => GetLayouts(ctx, masterIndex),
            "get_masters" => GetMasters(ctx),
            "apply_master" => ApplyMaster(ctx, outputPath, slideIndices, masterIndex, layoutIndex),
            "apply_layout_range" => ApplyLayoutRange(ctx, outputPath, slideIndices, layout),
            "apply_theme" => ApplyTheme(ctx, outputPath, themePath),
            _ => throw new ArgumentException($"Unknown operation: {operation}")
        };
    }

    /// <summary>
    ///     Sets layout for a slide.
    /// </summary>
    /// <param name="ctx">The document context.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="slideIndex">The slide index (0-based).</param>
    /// <param name="layoutStr">The layout type string.</param>
    /// <returns>A message indicating the result of the operation.</returns>
    /// <exception cref="ArgumentException">Thrown when slideIndex or layout is not provided, or layout type is invalid.</exception>
    private static string SetLayout(DocumentContext<Presentation> ctx, string? outputPath, int? slideIndex,
        string? layoutStr)
    {
        if (!slideIndex.HasValue)
            throw new ArgumentException("slideIndex is required for set operation");
        if (string.IsNullOrEmpty(layoutStr))
            throw new ArgumentException("layout is required for set operation");

        var presentation = ctx.Document;
        PowerPointHelper.ValidateCollectionIndex(slideIndex.Value, presentation.Slides.Count, "slide");

        var layout = FindLayoutByType(presentation, layoutStr);
        presentation.Slides[slideIndex.Value].LayoutSlide = layout;

        ctx.Save(outputPath);

        var result = $"Layout '{layoutStr}' set for slide {slideIndex}. ";
        result += ctx.GetOutputMessage(outputPath);
        return result;
    }

    /// <summary>
    ///     Gets available layouts with layout type information.
    /// </summary>
    /// <param name="ctx">The document context.</param>
    /// <param name="masterIndex">The master slide index (optional).</param>
    /// <returns>A JSON string containing layout information.</returns>
    private static string GetLayouts(DocumentContext<Presentation> ctx, int? masterIndex)
    {
        var presentation = ctx.Document;

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
            List<object> mastersList = [];

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
    }

    /// <summary>
    ///     Gets master slides information.
    /// </summary>
    /// <param name="ctx">The document context.</param>
    /// <returns>A JSON string containing master slide information.</returns>
    private static string GetMasters(DocumentContext<Presentation> ctx)
    {
        var presentation = ctx.Document;

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

        List<object> mastersList = [];

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
    }

    /// <summary>
    ///     Applies a master slide layout to specified slides.
    /// </summary>
    /// <param name="ctx">The document context.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="slideIndicesJson">JSON array of slide indices.</param>
    /// <param name="masterIndex">The master slide index.</param>
    /// <param name="layoutIndex">The layout index under the master.</param>
    /// <returns>A message indicating the result of the operation.</returns>
    /// <exception cref="ArgumentException">Thrown when masterIndex or layoutIndex is not provided.</exception>
    private static string ApplyMaster(DocumentContext<Presentation> ctx, string? outputPath,
        string? slideIndicesJson, int? masterIndex, int? layoutIndex)
    {
        if (!masterIndex.HasValue)
            throw new ArgumentException("masterIndex is required for apply_master operation");
        if (!layoutIndex.HasValue)
            throw new ArgumentException("layoutIndex is required for apply_master operation");

        var presentation = ctx.Document;

        PowerPointHelper.ValidateCollectionIndex(masterIndex.Value, presentation.Masters.Count, "master");
        var master = presentation.Masters[masterIndex.Value];
        PowerPointHelper.ValidateCollectionIndex(layoutIndex.Value, master.LayoutSlides.Count, "layout");

        var slideIndicesArray = ParseSlideIndicesJson(slideIndicesJson);
        var targets = slideIndicesArray?.Length > 0
            ? slideIndicesArray
            : Enumerable.Range(0, presentation.Slides.Count).ToArray();

        ValidateSlideIndices(targets, presentation.Slides.Count);

        var layout = master.LayoutSlides[layoutIndex.Value];
        foreach (var idx in targets)
            presentation.Slides[idx].LayoutSlide = layout;

        ctx.Save(outputPath);

        var result = $"Master {masterIndex} / Layout {layoutIndex} applied to {targets.Length} slides. ";
        result += ctx.GetOutputMessage(outputPath);
        return result;
    }

    /// <summary>
    ///     Applies layout to a range of slides.
    /// </summary>
    /// <param name="ctx">The document context.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="slideIndicesJson">JSON array of slide indices.</param>
    /// <param name="layoutStr">The layout type string.</param>
    /// <returns>A message indicating the result of the operation.</returns>
    /// <exception cref="ArgumentException">Thrown when layout or slideIndices is not provided.</exception>
    private static string ApplyLayoutRange(DocumentContext<Presentation> ctx, string? outputPath,
        string? slideIndicesJson, string? layoutStr)
    {
        if (string.IsNullOrEmpty(layoutStr))
            throw new ArgumentException("layout is required for apply_layout_range operation");

        var slideIndicesArray = ParseSlideIndicesJson(slideIndicesJson);
        if (slideIndicesArray == null || slideIndicesArray.Length == 0)
            throw new ArgumentException("slideIndices is required for apply_layout_range operation");

        var presentation = ctx.Document;

        ValidateSlideIndices(slideIndicesArray, presentation.Slides.Count);

        var layout = FindLayoutByType(presentation, layoutStr);

        foreach (var idx in slideIndicesArray)
            presentation.Slides[idx].LayoutSlide = layout;

        ctx.Save(outputPath);

        var result = $"Layout '{layoutStr}' applied to {slideIndicesArray.Length} slide(s). ";
        result += ctx.GetOutputMessage(outputPath);
        return result;
    }

    /// <summary>
    ///     Applies a theme to the presentation by copying master slides.
    /// </summary>
    /// <param name="ctx">The document context.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="themePath">The theme template file path.</param>
    /// <returns>A message indicating the result of the operation.</returns>
    /// <exception cref="ArgumentException">Thrown when themePath is not provided.</exception>
    /// <exception cref="FileNotFoundException">Thrown when the theme file is not found.</exception>
    /// <exception cref="InvalidOperationException">Thrown when the theme presentation has no master slides.</exception>
    private static string ApplyTheme(DocumentContext<Presentation> ctx, string? outputPath, string? themePath)
    {
        if (string.IsNullOrEmpty(themePath))
            throw new ArgumentException("themePath is required for apply_theme operation");
        if (!File.Exists(themePath))
            throw new FileNotFoundException($"Theme file not found: {themePath}");

        var presentation = ctx.Document;
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

        ctx.Save(outputPath);

        var result = $"Theme applied ({copiedCount} master(s) copied, layout applied to all slides). ";
        result += ctx.GetOutputMessage(outputPath);
        return result;
    }

    #region Helper Methods

    /// <summary>
    ///     Parses JSON array of slide indices.
    /// </summary>
    /// <param name="slideIndicesJson">The JSON string containing slide indices array.</param>
    /// <returns>An array of slide indices, or null if input is empty.</returns>
    /// <exception cref="ArgumentException">Thrown when the JSON format is invalid.</exception>
    private static int[]? ParseSlideIndicesJson(string? slideIndicesJson)
    {
        if (string.IsNullOrWhiteSpace(slideIndicesJson))
            return null;

        try
        {
            return JsonSerializer.Deserialize<int[]>(slideIndicesJson);
        }
        catch (JsonException)
        {
            throw new ArgumentException(
                $"Invalid slideIndices format. Expected JSON array, e.g., [0,1,2]. Got: {slideIndicesJson}");
        }
    }

    /// <summary>
    ///     Finds a layout slide by layout type string.
    /// </summary>
    /// <param name="presentation">The presentation object.</param>
    /// <param name="layoutStr">The layout type string.</param>
    /// <returns>The matching layout slide.</returns>
    /// <exception cref="ArgumentException">Thrown when the layout type is unknown or not found in the presentation.</exception>
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
    /// <param name="indices">The array of slide indices to validate.</param>
    /// <param name="slideCount">The total number of slides in the presentation.</param>
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
    /// <param name="layoutSlides">The collection of layout slides.</param>
    /// <returns>A list of objects containing layout information.</returns>
    private static List<object> BuildLayoutsList(IMasterLayoutSlideCollection layoutSlides)
    {
        List<object> layoutsList = [];
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