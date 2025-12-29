using System.Text.Json;
using System.Text.Json.Nodes;
using Aspose.Slides;
using Aspose.Slides.Export;
using AsposeMcpServer.Core;

namespace AsposeMcpServer.Tools.PowerPoint;

/// <summary>
///     Unified tool for managing PowerPoint slides (add, delete, get info, move, duplicate, hide, clear, edit)
///     Merges: PptAddSlideTool, PptDeleteSlideTool, PptGetSlidesInfoTool, PptMoveSlideTool,
///     PptDuplicateSlideTool, PptHideSlidesTool, PptClearSlideTool, PptEditSlideTool
/// </summary>
public class PptSlideTool : IAsposeTool
{
    public string Description =>
        @"Manage PowerPoint slides. Supports 8 operations: add, delete, get_info, move, duplicate, hide, clear, edit.

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
                description = "Slide layout type (optional, for add operation)",
                @enum = new[] { "Blank", "Title", "TitleOnly", "TwoColumn", "SectionHeader" }
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
            },
            outputPath = new
            {
                type = "string",
                description =
                    "Output file path (optional, for add/copy/delete/edit/hide operations, defaults to input path)"
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
            "add" => await AddSlideAsync(path, outputPath, arguments),
            "delete" => await DeleteSlideAsync(path, outputPath, arguments),
            "get_info" => await GetSlidesInfoAsync(path),
            "move" => await MoveSlideAsync(path, outputPath, arguments),
            "duplicate" => await DuplicateSlideAsync(path, outputPath, arguments),
            "hide" => await HideSlidesAsync(path, outputPath, arguments),
            "clear" => await ClearSlideAsync(path, outputPath, arguments),
            "edit" => await EditSlideAsync(path, outputPath, arguments),
            _ => throw new ArgumentException($"Unknown operation: {operation}")
        };
    }

    /// <summary>
    ///     Adds a new slide to the presentation.
    /// </summary>
    /// <param name="path">PowerPoint file path.</param>
    /// <param name="outputPath">Output file path.</param>
    /// <param name="arguments">JSON arguments containing optional layoutType.</param>
    /// <returns>Success message with slide count.</returns>
    /// <exception cref="InvalidOperationException">Thrown when presentation has no layout slides.</exception>
    private Task<string> AddSlideAsync(string path, string outputPath, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var layoutTypeStr = ArgumentHelper.GetString(arguments, "layoutType", "Blank");

            using var presentation = new Presentation(path);

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

            presentation.Save(outputPath, SaveFormat.Pptx);

            return $"Slide added (total: {presentation.Slides.Count}). Output: {outputPath}";
        });
    }

    /// <summary>
    ///     Deletes a slide from the presentation.
    /// </summary>
    /// <param name="path">PowerPoint file path.</param>
    /// <param name="outputPath">Output file path.</param>
    /// <param name="arguments">JSON arguments containing slideIndex.</param>
    /// <returns>Success message.</returns>
    /// <exception cref="ArgumentException">Thrown when slideIndex is out of range.</exception>
    /// <exception cref="InvalidOperationException">Thrown when attempting to delete the last slide.</exception>
    private Task<string> DeleteSlideAsync(string path, string outputPath, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var slideIndex = ArgumentHelper.GetInt(arguments, "slideIndex");

            using var presentation = new Presentation(path);
            if (slideIndex < 0 || slideIndex >= presentation.Slides.Count)
                throw new ArgumentException($"slideIndex must be between 0 and {presentation.Slides.Count - 1}");

            if (presentation.Slides.Count == 1)
                throw new InvalidOperationException(
                    "Cannot delete the last slide. A presentation must have at least one slide.");

            presentation.Slides.RemoveAt(slideIndex);
            presentation.Save(outputPath, SaveFormat.Pptx);

            return $"Slide {slideIndex} deleted ({presentation.Slides.Count} remaining). Output: {outputPath}";
        });
    }

    /// <summary>
    ///     Gets information about all slides.
    /// </summary>
    /// <param name="path">PowerPoint file path.</param>
    /// <returns>JSON string with slide information including layout details.</returns>
    private Task<string> GetSlidesInfoAsync(string path)
    {
        return Task.Run(() =>
        {
            using var presentation = new Presentation(path);

            var slidesList = new List<object>();
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
        });
    }

    /// <summary>
    ///     Moves a slide to a different position using Reorder method.
    /// </summary>
    /// <param name="path">PowerPoint file path.</param>
    /// <param name="outputPath">Output file path.</param>
    /// <param name="arguments">JSON arguments containing fromIndex, toIndex.</param>
    /// <returns>Success message.</returns>
    /// <exception cref="ArgumentException">Thrown when fromIndex or toIndex is out of range.</exception>
    private Task<string> MoveSlideAsync(string path, string outputPath, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var fromIndex = ArgumentHelper.GetInt(arguments, "fromIndex");
            var toIndex = ArgumentHelper.GetInt(arguments, "toIndex");

            using var presentation = new Presentation(path);
            var count = presentation.Slides.Count;

            if (fromIndex < 0 || fromIndex >= count)
                throw new ArgumentException($"fromIndex must be between 0 and {count - 1}");
            if (toIndex < 0 || toIndex >= count)
                throw new ArgumentException($"toIndex must be between 0 and {count - 1}");

            var slide = presentation.Slides[fromIndex];
            presentation.Slides.Reorder(toIndex, slide);
            presentation.Save(outputPath, SaveFormat.Pptx);

            return $"Slide moved from {fromIndex} to {toIndex} (total: {count}). Output: {outputPath}";
        });
    }

    /// <summary>
    ///     Duplicates a slide.
    /// </summary>
    /// <param name="path">PowerPoint file path.</param>
    /// <param name="outputPath">Output file path.</param>
    /// <param name="arguments">JSON arguments containing slideIndex, optional insertAt.</param>
    /// <returns>Success message with new slide count.</returns>
    /// <exception cref="ArgumentException">Thrown when slideIndex or insertAt is out of range.</exception>
    private Task<string> DuplicateSlideAsync(string path, string outputPath, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var slideIndex = ArgumentHelper.GetInt(arguments, "slideIndex");
            var insertAt = ArgumentHelper.GetIntNullable(arguments, "insertAt");

            using var presentation = new Presentation(path);
            var count = presentation.Slides.Count;

            if (slideIndex < 0 || slideIndex >= count)
                throw new ArgumentException($"slideIndex must be between 0 and {count - 1}");

            if (insertAt.HasValue)
            {
                if (insertAt.Value < 0 || insertAt.Value > count)
                    throw new ArgumentException($"insertAt must be between 0 and {count}");

                presentation.Slides.InsertClone(insertAt.Value, presentation.Slides[slideIndex]);
            }
            else
            {
                presentation.Slides.AddClone(presentation.Slides[slideIndex]);
            }

            presentation.Save(outputPath, SaveFormat.Pptx);
            return $"Slide {slideIndex} duplicated (total: {presentation.Slides.Count}). Output: {outputPath}";
        });
    }

    /// <summary>
    ///     Hides or shows slides.
    /// </summary>
    /// <param name="path">PowerPoint file path.</param>
    /// <param name="outputPath">Output file path.</param>
    /// <param name="arguments">JSON arguments containing slideIndices array, hidden.</param>
    /// <returns>Success message.</returns>
    /// <exception cref="ArgumentException">Thrown when any slide index is out of range.</exception>
    private Task<string> HideSlidesAsync(string path, string outputPath, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var hidden = ArgumentHelper.GetBool(arguments, "hidden", false);
            var slideIndicesArray = ArgumentHelper.GetArray(arguments, "slideIndices", false);
            var slideIndices = slideIndicesArray?.Select(x => x?.GetValue<int>() ?? -1).ToArray();

            using var presentation = new Presentation(path);
            var targets = slideIndices?.Length > 0
                ? slideIndices
                : Enumerable.Range(0, presentation.Slides.Count).ToArray();

            foreach (var idx in targets)
                if (idx < 0 || idx >= presentation.Slides.Count)
                    throw new ArgumentException($"slide index {idx} out of range");

            foreach (var idx in targets) presentation.Slides[idx].Hidden = hidden;

            presentation.Save(outputPath, SaveFormat.Pptx);
            return $"Set {targets.Length} slide(s) hidden={hidden}. Output: {outputPath}";
        });
    }

    /// <summary>
    ///     Clears all content from a slide.
    /// </summary>
    /// <param name="path">PowerPoint file path.</param>
    /// <param name="outputPath">Output file path.</param>
    /// <param name="arguments">JSON arguments containing slideIndex.</param>
    /// <returns>Success message.</returns>
    /// <exception cref="ArgumentException">Thrown when slideIndex is out of range.</exception>
    private Task<string> ClearSlideAsync(string path, string outputPath, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var slideIndex = ArgumentHelper.GetInt(arguments, "slideIndex");

            using var presentation = new Presentation(path);
            var slide = PowerPointHelper.GetSlide(presentation, slideIndex);
            while (slide.Shapes.Count > 0) slide.Shapes.RemoveAt(0);

            presentation.Save(outputPath, SaveFormat.Pptx);
            return $"Cleared all shapes from slide {slideIndex}. Output: {outputPath}";
        });
    }

    /// <summary>
    ///     Edits slide properties.
    /// </summary>
    /// <param name="path">PowerPoint file path.</param>
    /// <param name="outputPath">Output file path.</param>
    /// <param name="arguments">JSON arguments containing slideIndex, optional layoutIndex.</param>
    /// <returns>Success message.</returns>
    /// <exception cref="ArgumentException">Thrown when slideIndex or layoutIndex is out of range.</exception>
    private Task<string> EditSlideAsync(string path, string outputPath, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var slideIndex = ArgumentHelper.GetInt(arguments, "slideIndex");
            var layoutIndex = ArgumentHelper.GetIntNullable(arguments, "layoutIndex");

            using var presentation = new Presentation(path);
            var slide = PowerPointHelper.GetSlide(presentation, slideIndex);

            if (layoutIndex.HasValue)
            {
                if (layoutIndex.Value < 0 || layoutIndex.Value >= presentation.LayoutSlides.Count)
                    throw new ArgumentException(
                        $"layoutIndex must be between 0 and {presentation.LayoutSlides.Count - 1}");
                slide.LayoutSlide = presentation.LayoutSlides[layoutIndex.Value];
            }

            presentation.Save(outputPath, SaveFormat.Pptx);
            return $"Slide {slideIndex} updated. Output: {outputPath}";
        });
    }
}