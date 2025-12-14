using System.Text.Json.Nodes;
using System.Text;
using System.Linq;
using Aspose.Slides;
using Aspose.Slides.Export;
using AsposeMcpServer.Core;

namespace AsposeMcpServer.Tools;

/// <summary>
/// Unified tool for managing PowerPoint slides (add, delete, get info, move, duplicate, hide, clear, edit)
/// Merges: PptAddSlideTool, PptDeleteSlideTool, PptGetSlidesInfoTool, PptMoveSlideTool, 
/// PptDuplicateSlideTool, PptHideSlidesTool, PptClearSlideTool, PptEditSlideTool
/// </summary>
public class PptSlideTool : IAsposeTool
{
    public string Description => @"Manage PowerPoint slides. Supports 8 operations: add, delete, get_info, move, duplicate, hide, clear, edit.

Usage examples:
- Add slide: ppt_slide(operation='add', path='presentation.pptx', layoutType='Blank')
- Delete slide: ppt_slide(operation='delete', path='presentation.pptx', slideIndex=0)
- Get info: ppt_slide(operation='get_info', path='presentation.pptx')
- Move slide: ppt_slide(operation='move', path='presentation.pptx', fromIndex=0, toIndex=2)
- Duplicate slide: ppt_slide(operation='duplicate', path='presentation.pptx', slideIndex=0)
- Hide slide: ppt_slide(operation='hide', path='presentation.pptx', slideIndex=0, hidden=true)
- Clear slide: ppt_slide(operation='clear', path='presentation.pptx', slideIndex=0)
- Edit slide: ppt_slide(operation='edit', path='presentation.pptx', slideIndex=0)";

    public object InputSchema => new
    {
        type = "object",
        properties = new
        {
            operation = new
            {
                type = "string",
                description = @"Operation to perform.
- 'add': Add a new slide (required params: path)
- 'delete': Delete a slide (required params: path, slideIndex)
- 'get_info': Get slides info (required params: path)
- 'move': Move a slide (required params: path, fromIndex, toIndex)
- 'duplicate': Duplicate a slide (required params: path, slideIndex)
- 'hide': Hide/show a slide (required params: path, slideIndex, hidden)
- 'clear': Clear slide content (required params: path, slideIndex)
- 'edit': Edit slide properties (required params: path, slideIndex)",
                @enum = new[] { "add", "delete", "get_info", "move", "duplicate", "hide", "clear", "edit" }
            },
            path = new
            {
                type = "string",
                description = "Presentation file path (required for all operations)"
            },
            slideIndex = new
            {
                type = "number",
                description = "Slide index (0-based, required for most operations)"
            },
            layoutType = new
            {
                type = "string",
                description = "Slide layout type (Blank, Title, TitleOnly, etc., optional, for add operation)"
            },
            fromIndex = new
            {
                type = "number",
                description = "Source slide index (0-based, required for move operation)"
            },
            toIndex = new
            {
                type = "number",
                description = "Target slide index (0-based, required for move operation)"
            },
            insertAt = new
            {
                type = "number",
                description = "Target index to insert clone (0-based, optional, for duplicate, default: append)"
            },
            slideIndices = new
            {
                type = "array",
                items = new { type = "number" },
                description = "Slide indices array (optional, for hide operation, if not provided affects all slides)"
            },
            hidden = new
            {
                type = "boolean",
                description = "Hide slides (true) or show (false, required for hide operation)"
            },
            layoutIndex = new
            {
                type = "number",
                description = "Layout index (0-based, optional, for edit operation)"
            }
        },
        required = new[] { "operation", "path" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var operation = ArgumentHelper.GetString(arguments, "operation", "operation");
        var path = ArgumentHelper.GetAndValidatePath(arguments);

        return operation.ToLower() switch
        {
            "add" => await AddSlideAsync(arguments, path),
            "delete" => await DeleteSlideAsync(arguments, path),
            "get_info" => await GetSlidesInfoAsync(arguments, path),
            "move" => await MoveSlideAsync(arguments, path),
            "duplicate" => await DuplicateSlideAsync(arguments, path),
            "hide" => await HideSlidesAsync(arguments, path),
            "clear" => await ClearSlideAsync(arguments, path),
            "edit" => await EditSlideAsync(arguments, path),
            _ => throw new ArgumentException($"Unknown operation: {operation}")
        };
    }

    /// <summary>
    /// Adds a new slide to the presentation
    /// </summary>
    /// <param name="arguments">JSON arguments containing optional layoutType, outputPath</param>
    /// <param name="path">PowerPoint file path</param>
    /// <returns>Success message with slide index</returns>
    private async Task<string> AddSlideAsync(JsonObject? arguments, string path)
    {
        var layoutType = arguments?["layoutType"]?.GetValue<string>() ?? "Blank";

        using var presentation = new Presentation(path);
        
        if (presentation.LayoutSlides.Count == 0)
        {
            throw new InvalidOperationException("Presentation has no layout slides");
        }
        
        var layoutSlide = presentation.LayoutSlides[0];
        var slide = presentation.Slides.AddEmptySlide(layoutSlide);

        presentation.Save(path, SaveFormat.Pptx);

        return await Task.FromResult($"Slide added to presentation: {path} (Total: {presentation.Slides.Count})");
    }

    /// <summary>
    /// Deletes a slide from the presentation
    /// </summary>
    /// <param name="arguments">JSON arguments containing slideIndex, optional outputPath</param>
    /// <param name="path">PowerPoint file path</param>
    /// <returns>Success message</returns>
    private async Task<string> DeleteSlideAsync(JsonObject? arguments, string path)
    {
        var slideIndex = ArgumentHelper.GetInt(arguments, "slideIndex", "slideIndex");

        using var presentation = new Presentation(path);
        if (slideIndex < 0 || slideIndex >= presentation.Slides.Count)
        {
            throw new ArgumentException($"slideIndex must be between 0 and {presentation.Slides.Count - 1}");
        }

        presentation.Slides.RemoveAt(slideIndex);
        presentation.Save(path, SaveFormat.Pptx);

        return await Task.FromResult($"已刪除投影片 {slideIndex}，剩餘 {presentation.Slides.Count} 張: {path}");
    }

    /// <summary>
    /// Gets information about all slides
    /// </summary>
    /// <param name="arguments">JSON arguments (no specific parameters required)</param>
    /// <param name="path">PowerPoint file path</param>
    /// <returns>Formatted string with slide information</returns>
    private async Task<string> GetSlidesInfoAsync(JsonObject? arguments, string path)
    {
        using var presentation = new Presentation(path);
        var sb = new StringBuilder();

        sb.AppendLine($"總投影片數: {presentation.Slides.Count}");

        for (int i = 0; i < presentation.Slides.Count; i++)
        {
            var slide = presentation.Slides[i];
            var title = slide.Shapes.FirstOrDefault(s => s.Placeholder?.Type == PlaceholderType.Title) as IAutoShape;
            var titleText = title?.TextFrame?.Text ?? "(無標題)";
            var notes = slide.NotesSlideManager.NotesSlide?.NotesTextFrame?.Text;

            sb.AppendLine($"\n--- 投影片 {i} ---");
            sb.AppendLine($"標題: {titleText}");
            sb.AppendLine($"形狀數: {slide.Shapes.Count}");
            sb.AppendLine($"是否有講者備註: {(string.IsNullOrWhiteSpace(notes) ? "否" : "是")}");
            sb.AppendLine($"隱藏: {slide.Hidden}");
        }

        return await Task.FromResult(sb.ToString());
    }

    /// <summary>
    /// Moves a slide to a different position
    /// </summary>
    /// <param name="arguments">JSON arguments containing slideIndex, targetIndex, optional outputPath</param>
    /// <param name="path">PowerPoint file path</param>
    /// <returns>Success message</returns>
    private async Task<string> MoveSlideAsync(JsonObject? arguments, string path)
    {
        var fromIndex = ArgumentHelper.GetInt(arguments, "fromIndex", "fromIndex");
        var toIndex = ArgumentHelper.GetInt(arguments, "toIndex", "toIndex");

        using var presentation = new Presentation(path);
        var count = presentation.Slides.Count;

        if (fromIndex < 0 || fromIndex >= count)
        {
            throw new ArgumentException($"fromIndex must be between 0 and {count - 1}");
        }
        if (toIndex < 0 || toIndex >= count)
        {
            throw new ArgumentException($"toIndex must be between 0 and {count - 1}");
        }

        var source = presentation.Slides[fromIndex];
        presentation.Slides.InsertClone(toIndex, source);
        var removeIndex = fromIndex + (fromIndex < toIndex ? 1 : 0);
        presentation.Slides.RemoveAt(removeIndex);
        presentation.Save(path, SaveFormat.Pptx);

        return await Task.FromResult($"投影片已移動: {fromIndex} -> {toIndex} (總數 {count})");
    }

    /// <summary>
    /// Duplicates a slide
    /// </summary>
    /// <param name="arguments">JSON arguments containing slideIndex, optional outputPath</param>
    /// <param name="path">PowerPoint file path</param>
    /// <returns>Success message with new slide index</returns>
    private async Task<string> DuplicateSlideAsync(JsonObject? arguments, string path)
    {
        var slideIndex = ArgumentHelper.GetInt(arguments, "slideIndex", "slideIndex");
        var insertAt = arguments?["insertAt"]?.GetValue<int?>();

        using var presentation = new Presentation(path);
        var count = presentation.Slides.Count;

        if (slideIndex < 0 || slideIndex >= count)
        {
            throw new ArgumentException($"slideIndex must be between 0 and {count - 1}");
        }

        if (insertAt.HasValue)
        {
            if (insertAt.Value < 0 || insertAt.Value > count)
            {
                throw new ArgumentException($"insertAt must be between 0 and {count}");
            }

            presentation.Slides.InsertClone(insertAt.Value, presentation.Slides[slideIndex]);
        }
        else
        {
            presentation.Slides.AddClone(presentation.Slides[slideIndex]);
        }

        presentation.Save(path, SaveFormat.Pptx);
        return await Task.FromResult($"已複製投影片 {slideIndex}，總數 {presentation.Slides.Count} 張: {path}");
    }

    /// <summary>
    /// Hides or shows slides
    /// </summary>
    /// <param name="arguments">JSON arguments containing slideIndexes array, isHidden, optional outputPath</param>
    /// <param name="path">PowerPoint file path</param>
    /// <returns>Success message</returns>
    private async Task<string> HideSlidesAsync(JsonObject? arguments, string path)
    {
        var hidden = arguments?["hidden"]?.GetValue<bool?>() ?? false;
        var slideIndices = arguments?["slideIndices"]?.AsArray()?.Select(x => x?.GetValue<int>() ?? -1).ToArray();

        using var presentation = new Presentation(path);
        var targets = slideIndices?.Length > 0
            ? slideIndices
            : Enumerable.Range(0, presentation.Slides.Count).ToArray();

        foreach (var idx in targets)
        {
            if (idx < 0 || idx >= presentation.Slides.Count)
            {
                throw new ArgumentException($"slide index {idx} out of range");
            }
        }

        foreach (var idx in targets)
        {
            presentation.Slides[idx].Hidden = hidden;
        }

        presentation.Save(path, SaveFormat.Pptx);
        return await Task.FromResult($"已設定 {targets.Length} 張投影片 Hidden={hidden}");
    }

    /// <summary>
    /// Clears all content from a slide
    /// </summary>
    /// <param name="arguments">JSON arguments containing slideIndex, optional outputPath</param>
    /// <param name="path">PowerPoint file path</param>
    /// <returns>Success message</returns>
    private async Task<string> ClearSlideAsync(JsonObject? arguments, string path)
    {
        var slideIndex = ArgumentHelper.GetInt(arguments, "slideIndex", "slideIndex");

        using var presentation = new Presentation(path);
        var slide = PowerPointHelper.GetSlide(presentation, slideIndex);
        while (slide.Shapes.Count > 0)
        {
            slide.Shapes.RemoveAt(0);
        }

        presentation.Save(path, SaveFormat.Pptx);
        return await Task.FromResult($"已清除投影片 {slideIndex} 的所有形狀");
    }

    /// <summary>
    /// Edits slide properties
    /// </summary>
    /// <param name="arguments">JSON arguments containing slideIndex, optional layoutType, background, outputPath</param>
    /// <param name="path">PowerPoint file path</param>
    /// <returns>Success message</returns>
    private async Task<string> EditSlideAsync(JsonObject? arguments, string path)
    {
        var slideIndex = ArgumentHelper.GetInt(arguments, "slideIndex", "slideIndex");
        var layoutIndex = arguments?["layoutIndex"]?.GetValue<int?>();

        using var presentation = new Presentation(path);
        var slide = PowerPointHelper.GetSlide(presentation, slideIndex);

        if (layoutIndex.HasValue)
        {
            if (layoutIndex.Value < 0 || layoutIndex.Value >= presentation.LayoutSlides.Count)
            {
                throw new ArgumentException($"layoutIndex must be between 0 and {presentation.LayoutSlides.Count - 1}");
            }
            slide.LayoutSlide = presentation.LayoutSlides[layoutIndex.Value];
        }

        presentation.Save(path, SaveFormat.Pptx);
        return await Task.FromResult($"Slide {slideIndex} updated: {path}");
    }
}

