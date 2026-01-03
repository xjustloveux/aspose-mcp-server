using System.ComponentModel;
using System.Text.Json;
using Aspose.Slides;
using AsposeMcpServer.Core.Helpers;
using AsposeMcpServer.Core.Session;
using ModelContextProtocol.Server;

namespace AsposeMcpServer.Tools.PowerPoint;

/// <summary>
///     Unified tool for managing PowerPoint slides (add, delete, get info, move, duplicate, hide, clear, edit)
///     Merges: PptAddSlideTool, PptDeleteSlideTool, PptGetSlidesInfoTool, PptMoveSlideTool,
///     PptDuplicateSlideTool, PptHideSlidesTool, PptClearSlideTool, PptEditSlideTool
/// </summary>
[McpServerToolType]
public class PptSlideTool
{
    /// <summary>
    ///     Session manager for document session handling.
    /// </summary>
    private readonly DocumentSessionManager? _sessionManager;

    /// <summary>
    ///     Initializes a new instance of the <see cref="PptSlideTool" /> class.
    /// </summary>
    /// <param name="sessionManager">Optional session manager for in-memory editing.</param>
    public PptSlideTool(DocumentSessionManager? sessionManager = null)
    {
        _sessionManager = sessionManager;
    }

    [McpServerTool(Name = "ppt_slide")]
    [Description(
        @"Manage PowerPoint slides. Supports 8 operations: add, delete, get_info, move, duplicate, hide, clear, edit.

Usage examples:
- Add slide: ppt_slide(operation='add', path='presentation.pptx', layoutType='Blank')
- Delete slide: ppt_slide(operation='delete', path='presentation.pptx', slideIndex=0)
- Get info: ppt_slide(operation='get_info', path='presentation.pptx')
- Move slide: ppt_slide(operation='move', path='presentation.pptx', fromIndex=0, toIndex=2)
- Duplicate slide: ppt_slide(operation='duplicate', path='presentation.pptx', slideIndex=0)
- Hide slide: ppt_slide(operation='hide', path='presentation.pptx', slideIndex=0, hidden=true)
- Clear slide: ppt_slide(operation='clear', path='presentation.pptx', slideIndex=0)
- Edit slide: ppt_slide(operation='edit', path='presentation.pptx', slideIndex=0)")]
    public string Execute(
        [Description(@"Operation to perform.
- 'add': Add a new slide (required params: path)
- 'delete': Delete a slide (required params: path, slideIndex)
- 'get_info': Get slides info (required params: path)
- 'move': Move a slide (required params: path, fromIndex, toIndex)
- 'duplicate': Duplicate a slide (required params: path, slideIndex)
- 'hide': Hide/show a slide (required params: path, slideIndex, hidden)
- 'clear': Clear slide content (required params: path, slideIndex)
- 'edit': Edit slide properties (required params: path, slideIndex)")]
        string operation,
        [Description("Presentation file path (required if no sessionId)")]
        string? path = null,
        [Description("Session ID for in-memory editing")]
        string? sessionId = null,
        [Description("Output file path (optional, defaults to input path)")]
        string? outputPath = null,
        [Description("Slide index (0-based, required for most operations)")]
        int? slideIndex = null,
        [Description("Slide layout type: Blank, Title, TitleOnly, TwoColumn, SectionHeader")]
        string layoutType = "Blank",
        [Description("Source slide index (0-based, required for move operation)")]
        int? fromIndex = null,
        [Description("Target slide index (0-based, required for move operation)")]
        int? toIndex = null,
        [Description("Target index to insert clone (0-based, optional, for duplicate, default: append)")]
        int? insertAt = null,
        [Description("Slide indices array as JSON (optional, for hide operation, if not provided affects all slides)")]
        string? slideIndices = null,
        [Description("Hide slides (true) or show (false, required for hide operation)")]
        bool hidden = false,
        [Description("Layout index (0-based, optional, for edit operation)")]
        int? layoutIndex = null)
    {
        using var ctx = DocumentContext<Presentation>.Create(_sessionManager, sessionId, path);

        return operation.ToLower() switch
        {
            "add" => AddSlide(ctx, outputPath, layoutType),
            "delete" => DeleteSlide(ctx, outputPath, slideIndex),
            "get_info" => GetSlidesInfo(ctx),
            "move" => MoveSlide(ctx, outputPath, fromIndex, toIndex),
            "duplicate" => DuplicateSlide(ctx, outputPath, slideIndex, insertAt),
            "hide" => HideSlides(ctx, outputPath, slideIndices, hidden),
            "clear" => ClearSlide(ctx, outputPath, slideIndex),
            "edit" => EditSlide(ctx, outputPath, slideIndex, layoutIndex),
            _ => throw new ArgumentException($"Unknown operation: {operation}")
        };
    }

    /// <summary>
    ///     Adds a new slide to the presentation.
    /// </summary>
    /// <param name="ctx">The document context.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="layoutTypeStr">The layout type string.</param>
    /// <returns>A message indicating the result of the operation.</returns>
    /// <exception cref="InvalidOperationException">Thrown when the presentation has no layout slides.</exception>
    private static string AddSlide(DocumentContext<Presentation> ctx, string? outputPath, string layoutTypeStr)
    {
        var presentation = ctx.Document;

        if (presentation.LayoutSlides.Count == 0)
            throw new InvalidOperationException("Presentation has no layout slides");

        var layoutType = layoutTypeStr.ToLower() switch
        {
            "title" => SlideLayoutType.Title,
            "titleonly" => SlideLayoutType.TitleOnly,
            "blank" => SlideLayoutType.Blank,
            "twocolumn" => SlideLayoutType.TwoColumnText,
            "sectionheader" => SlideLayoutType.SectionHeader,
            _ => SlideLayoutType.Custom
        };

        var layoutSlide = presentation.LayoutSlides.FirstOrDefault(ls => ls.LayoutType == layoutType) ??
                          presentation.LayoutSlides[0];
        _ = presentation.Slides.AddEmptySlide(layoutSlide);

        ctx.Save(outputPath);

        return $"Slide added (total: {presentation.Slides.Count}). {ctx.GetOutputMessage(outputPath)}";
    }

    /// <summary>
    ///     Deletes a slide from the presentation.
    /// </summary>
    /// <param name="ctx">The document context.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="slideIndex">The zero-based index of the slide to delete.</param>
    /// <returns>A message indicating the result of the operation.</returns>
    /// <exception cref="ArgumentException">Thrown when slideIndex is not provided or out of range.</exception>
    /// <exception cref="InvalidOperationException">Thrown when attempting to delete the last slide.</exception>
    private static string DeleteSlide(DocumentContext<Presentation> ctx, string? outputPath, int? slideIndex)
    {
        if (!slideIndex.HasValue)
            throw new ArgumentException("slideIndex is required for delete operation");

        var presentation = ctx.Document;
        if (slideIndex.Value < 0 || slideIndex.Value >= presentation.Slides.Count)
            throw new ArgumentException($"slideIndex must be between 0 and {presentation.Slides.Count - 1}");

        if (presentation.Slides.Count == 1)
            throw new InvalidOperationException(
                "Cannot delete the last slide. A presentation must have at least one slide.");

        presentation.Slides.RemoveAt(slideIndex.Value);
        ctx.Save(outputPath);

        return
            $"Slide {slideIndex.Value} deleted ({presentation.Slides.Count} remaining). {ctx.GetOutputMessage(outputPath)}";
    }

    /// <summary>
    ///     Gets information about all slides.
    /// </summary>
    /// <param name="ctx">The document context.</param>
    /// <returns>A JSON string containing slides information including count, slide details, and available layouts.</returns>
    private static string GetSlidesInfo(DocumentContext<Presentation> ctx)
    {
        var presentation = ctx.Document;

        List<object> slidesList = [];
        for (var i = 0; i < presentation.Slides.Count; i++)
        {
            var slide = presentation.Slides[i];
            var title = slide.Shapes.FirstOrDefault(s =>
                s.Placeholder?.Type == PlaceholderType.Title) as IAutoShape;
            var titleText = title?.TextFrame?.Text ?? "(no title)";
            var notes = slide.NotesSlideManager.NotesSlide?.NotesTextFrame?.Text;

            slidesList.Add(new
            {
                index = i,
                title = titleText,
                layoutType = slide.LayoutSlide.LayoutType.ToString(),
                layoutName = slide.LayoutSlide.Name,
                shapesCount = slide.Shapes.Count,
                hasSpeakerNotes = !string.IsNullOrWhiteSpace(notes),
                hidden = slide.Hidden
            });
        }

        var layoutsList = presentation.LayoutSlides
            .Select((ls, idx) => new { index = idx, name = ls.Name, type = ls.LayoutType.ToString() })
            .ToList();

        var result = new
        {
            count = presentation.Slides.Count,
            slides = slidesList,
            availableLayouts = layoutsList
        };

        return JsonSerializer.Serialize(result, new JsonSerializerOptions { WriteIndented = true });
    }

    /// <summary>
    ///     Moves a slide to a different position using Reorder method.
    /// </summary>
    /// <param name="ctx">The document context.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="fromIndex">The source slide index (0-based).</param>
    /// <param name="toIndex">The target slide index (0-based).</param>
    /// <returns>A message indicating the result of the operation.</returns>
    /// <exception cref="ArgumentException">Thrown when fromIndex or toIndex is not provided or out of range.</exception>
    private static string MoveSlide(DocumentContext<Presentation> ctx, string? outputPath, int? fromIndex, int? toIndex)
    {
        if (!fromIndex.HasValue)
            throw new ArgumentException("fromIndex is required for move operation");
        if (!toIndex.HasValue)
            throw new ArgumentException("toIndex is required for move operation");

        var presentation = ctx.Document;
        var count = presentation.Slides.Count;

        if (fromIndex.Value < 0 || fromIndex.Value >= count)
            throw new ArgumentException($"fromIndex must be between 0 and {count - 1}");
        if (toIndex.Value < 0 || toIndex.Value >= count)
            throw new ArgumentException($"toIndex must be between 0 and {count - 1}");

        var slide = presentation.Slides[fromIndex.Value];
        presentation.Slides.Reorder(toIndex.Value, slide);
        ctx.Save(outputPath);

        return
            $"Slide moved from {fromIndex.Value} to {toIndex.Value} (total: {count}). {ctx.GetOutputMessage(outputPath)}";
    }

    /// <summary>
    ///     Duplicates a slide.
    /// </summary>
    /// <param name="ctx">The document context.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="slideIndex">The zero-based index of the slide to duplicate.</param>
    /// <param name="insertAt">The target index to insert the clone (0-based, optional).</param>
    /// <returns>A message indicating the result of the operation.</returns>
    /// <exception cref="ArgumentException">Thrown when slideIndex is not provided or out of range.</exception>
    private static string DuplicateSlide(DocumentContext<Presentation> ctx, string? outputPath, int? slideIndex,
        int? insertAt)
    {
        if (!slideIndex.HasValue)
            throw new ArgumentException("slideIndex is required for duplicate operation");

        var presentation = ctx.Document;
        var count = presentation.Slides.Count;

        if (slideIndex.Value < 0 || slideIndex.Value >= count)
            throw new ArgumentException($"slideIndex must be between 0 and {count - 1}");

        if (insertAt.HasValue)
        {
            if (insertAt.Value < 0 || insertAt.Value > count)
                throw new ArgumentException($"insertAt must be between 0 and {count}");

            presentation.Slides.InsertClone(insertAt.Value, presentation.Slides[slideIndex.Value]);
        }
        else
        {
            presentation.Slides.AddClone(presentation.Slides[slideIndex.Value]);
        }

        ctx.Save(outputPath);
        return
            $"Slide {slideIndex.Value} duplicated (total: {presentation.Slides.Count}). {ctx.GetOutputMessage(outputPath)}";
    }

    /// <summary>
    ///     Hides or shows slides.
    /// </summary>
    /// <param name="ctx">The document context.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="slideIndicesJson">JSON array of slide indices to hide/show, or null to affect all slides.</param>
    /// <param name="hidden">True to hide slides, false to show them.</param>
    /// <returns>A message indicating the result of the operation.</returns>
    /// <exception cref="ArgumentException">Thrown when a slide index is out of range.</exception>
    private static string HideSlides(DocumentContext<Presentation> ctx, string? outputPath, string? slideIndicesJson,
        bool hidden)
    {
        var presentation = ctx.Document;

        int[] targets;
        if (!string.IsNullOrWhiteSpace(slideIndicesJson))
        {
            var indices = JsonSerializer.Deserialize<int[]>(slideIndicesJson);
            targets = indices ?? Enumerable.Range(0, presentation.Slides.Count).ToArray();
        }
        else
        {
            targets = Enumerable.Range(0, presentation.Slides.Count).ToArray();
        }

        foreach (var idx in targets)
            if (idx < 0 || idx >= presentation.Slides.Count)
                throw new ArgumentException($"slide index {idx} out of range");

        foreach (var idx in targets) presentation.Slides[idx].Hidden = hidden;

        ctx.Save(outputPath);
        return $"Set {targets.Length} slide(s) hidden={hidden}. {ctx.GetOutputMessage(outputPath)}";
    }

    /// <summary>
    ///     Clears all content from a slide.
    /// </summary>
    /// <param name="ctx">The document context.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="slideIndex">The zero-based index of the slide to clear.</param>
    /// <returns>A message indicating the result of the operation.</returns>
    /// <exception cref="ArgumentException">Thrown when slideIndex is not provided or out of range.</exception>
    private static string ClearSlide(DocumentContext<Presentation> ctx, string? outputPath, int? slideIndex)
    {
        if (!slideIndex.HasValue)
            throw new ArgumentException("slideIndex is required for clear operation");

        var presentation = ctx.Document;
        var slide = PowerPointHelper.GetSlide(presentation, slideIndex.Value);
        while (slide.Shapes.Count > 0) slide.Shapes.RemoveAt(0);

        ctx.Save(outputPath);
        return $"Cleared all shapes from slide {slideIndex.Value}. {ctx.GetOutputMessage(outputPath)}";
    }

    /// <summary>
    ///     Edits slide properties.
    /// </summary>
    /// <param name="ctx">The document context.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="slideIndex">The zero-based index of the slide to edit.</param>
    /// <param name="layoutIndex">The zero-based index of the layout to apply (optional).</param>
    /// <returns>A message indicating the result of the operation.</returns>
    /// <exception cref="ArgumentException">Thrown when slideIndex is not provided or layoutIndex is out of range.</exception>
    private static string EditSlide(DocumentContext<Presentation> ctx, string? outputPath, int? slideIndex,
        int? layoutIndex)
    {
        if (!slideIndex.HasValue)
            throw new ArgumentException("slideIndex is required for edit operation");

        var presentation = ctx.Document;
        var slide = PowerPointHelper.GetSlide(presentation, slideIndex.Value);

        if (layoutIndex.HasValue)
        {
            if (layoutIndex.Value < 0 || layoutIndex.Value >= presentation.LayoutSlides.Count)
                throw new ArgumentException(
                    $"layoutIndex must be between 0 and {presentation.LayoutSlides.Count - 1}");
            slide.LayoutSlide = presentation.LayoutSlides[layoutIndex.Value];
        }

        ctx.Save(outputPath);
        return $"Slide {slideIndex.Value} updated. {ctx.GetOutputMessage(outputPath)}";
    }
}