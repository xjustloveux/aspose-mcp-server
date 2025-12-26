using System.Text.Json;
using System.Text.Json.Nodes;
using Aspose.Slides;
using Aspose.Slides.Export;
using AsposeMcpServer.Core;

namespace AsposeMcpServer.Tools.PowerPoint;

/// <summary>
///     Unified tool for managing PowerPoint notes (add, edit, get, clear)
///     Merges: PptAddNotesTool, PptEditNotesTool, PptGetNotesTool, PptClearNotesTool
/// </summary>
public class PptNotesTool : IAsposeTool
{
    public string Description => @"Manage PowerPoint notes. Supports 4 operations: add, edit, get, clear.

Usage examples:
- Add notes: ppt_notes(operation='add', path='presentation.pptx', slideIndex=0, notes='Speaker notes')
- Edit notes: ppt_notes(operation='edit', path='presentation.pptx', slideIndex=0, notes='Updated notes')
- Get notes: ppt_notes(operation='get', path='presentation.pptx', slideIndex=0)
- Clear notes: ppt_notes(operation='clear', path='presentation.pptx', slideIndices=[0,1,2])";

    public object InputSchema => new
    {
        type = "object",
        properties = new
        {
            operation = new
            {
                type = "string",
                description = @"Operation to perform.
- 'add': Add notes to slide (required params: path, slideIndex, notes)
- 'edit': Edit slide notes (required params: path, slideIndex, notes)
- 'get': Get slide notes (required params: path, slideIndex)
- 'clear': Clear notes (required params: path)",
                @enum = new[] { "add", "edit", "get", "clear" }
            },
            path = new
            {
                type = "string",
                description = "Presentation file path (required for all operations)"
            },
            slideIndex = new
            {
                type = "number",
                description = "Slide index (0-based, required for add/edit, optional for get/clear)"
            },
            notes = new
            {
                type = "string",
                description = "Notes text content (required for add/edit)"
            },
            slideIndices = new
            {
                type = "array",
                items = new { type = "number" },
                description = "Slide indices array (optional, for clear, if not provided affects all slides)"
            },
            outputPath = new
            {
                type = "string",
                description = "Output file path (optional, for set/clear operations, defaults to input path)"
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
        var outputPath = ArgumentHelper.GetAndValidateOutputPath(arguments, path);

        return operation.ToLower() switch
        {
            "add" => await AddNotesAsync(path, outputPath, arguments),
            "edit" => await EditNotesAsync(path, outputPath, arguments),
            "get" => await GetNotesAsync(path, arguments),
            "clear" => await ClearNotesAsync(path, outputPath, arguments),
            _ => throw new ArgumentException($"Unknown operation: {operation}")
        };
    }

    /// <summary>
    ///     Adds notes to a slide
    /// </summary>
    /// <param name="path">PowerPoint file path</param>
    /// <param name="outputPath">Output file path</param>
    /// <param name="arguments">JSON arguments containing slideIndex, notesText</param>
    /// <returns>Success message</returns>
    private Task<string> AddNotesAsync(string path, string outputPath, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var slideIndex = ArgumentHelper.GetInt(arguments, "slideIndex");
            var notes = ArgumentHelper.GetString(arguments, "notes");

            using var presentation = new Presentation(path);
            var slide = PowerPointHelper.GetSlide(presentation, slideIndex);
            var notesSlide = slide.NotesSlideManager.NotesSlide ?? slide.NotesSlideManager.AddNotesSlide();
            var textFrame = notesSlide.NotesTextFrame;
            if (textFrame == null)
                throw new InvalidOperationException(
                    "Unable to get NotesTextFrame, file may be corrupted or format not supported");
            textFrame.Paragraphs.Clear();
            var para = new Paragraph();
            para.Portions.Add(new Portion(notes));
            textFrame.Paragraphs.Add(para);

            presentation.Save(outputPath, SaveFormat.Pptx);
            return $"Speaker notes added to slide {slideIndex}. Output: {outputPath}";
        });
    }

    /// <summary>
    ///     Edits slide notes
    /// </summary>
    /// <param name="path">PowerPoint file path</param>
    /// <param name="outputPath">Output file path</param>
    /// <param name="arguments">JSON arguments containing slideIndex, notesText</param>
    /// <returns>Success message</returns>
    private Task<string> EditNotesAsync(string path, string outputPath, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var slideIndex = ArgumentHelper.GetInt(arguments, "slideIndex");
            var notes = ArgumentHelper.GetString(arguments, "notes");

            using var presentation = new Presentation(path);
            var slide = PowerPointHelper.GetSlide(presentation, slideIndex);
            var notesSlide = slide.NotesSlideManager.NotesSlide ?? slide.NotesSlideManager.AddNotesSlide();
            notesSlide.NotesTextFrame.Text = notes;

            presentation.Save(outputPath, SaveFormat.Pptx);

            return $"Notes updated for slide {slideIndex}. Output: {outputPath}";
        });
    }

    /// <summary>
    ///     Gets notes from a slide
    /// </summary>
    /// <param name="path">PowerPoint file path</param>
    /// <param name="arguments">JSON arguments containing optional slideIndex (if null, gets all)</param>
    /// <returns>JSON string with notes</returns>
    private Task<string> GetNotesAsync(string path, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var slideIndex = ArgumentHelper.GetIntNullable(arguments, "slideIndex");

            using var presentation = new Presentation(path);

            if (slideIndex.HasValue)
            {
                if (slideIndex.Value < 0 || slideIndex.Value >= presentation.Slides.Count)
                    throw new ArgumentException($"slideIndex must be between 0 and {presentation.Slides.Count - 1}");

                var notesSlide = presentation.Slides[slideIndex.Value].NotesSlideManager.NotesSlide;
                var notesText = notesSlide?.NotesTextFrame?.Text;

                var result = new
                {
                    slideIndex = slideIndex.Value,
                    hasNotes = !string.IsNullOrWhiteSpace(notesText),
                    notes = notesText
                };

                return JsonSerializer.Serialize(result, new JsonSerializerOptions { WriteIndented = true });
            }
            else
            {
                var notesList = new List<object>();
                for (var i = 0; i < presentation.Slides.Count; i++)
                {
                    var notesSlide = presentation.Slides[i].NotesSlideManager.NotesSlide;
                    var notesText = notesSlide?.NotesTextFrame?.Text;

                    notesList.Add(new
                    {
                        slideIndex = i,
                        hasNotes = !string.IsNullOrWhiteSpace(notesText),
                        notes = notesText
                    });
                }

                var result = new
                {
                    count = presentation.Slides.Count,
                    slides = notesList
                };

                return JsonSerializer.Serialize(result, new JsonSerializerOptions { WriteIndented = true });
            }
        });
    }

    /// <summary>
    ///     Clears notes from a slide
    /// </summary>
    /// <param name="path">PowerPoint file path</param>
    /// <param name="outputPath">Output file path</param>
    /// <param name="arguments">JSON arguments containing slideIndex</param>
    /// <returns>Success message</returns>
    private Task<string> ClearNotesAsync(string path, string outputPath, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var slideIndicesArray = ArgumentHelper.GetArray(arguments, "slideIndices", false);
            var slideIndices = slideIndicesArray?.Select(x => x?.GetValue<int>() ?? -1).ToArray();

            using var presentation = new Presentation(path);
            var targets = slideIndices?.Length > 0
                ? slideIndices
                : Enumerable.Range(0, presentation.Slides.Count).ToArray();

            foreach (var idx in targets)
                if (idx < 0 || idx >= presentation.Slides.Count)
                    throw new ArgumentException($"slide index {idx} out of range");

            foreach (var idx in targets)
            {
                var slide = presentation.Slides[idx];
                var notesSlide = slide.NotesSlideManager.NotesSlide ?? slide.NotesSlideManager.AddNotesSlide();
                if (notesSlide.NotesTextFrame != null) notesSlide.NotesTextFrame.Text = string.Empty;
            }

            presentation.Save(outputPath, SaveFormat.Pptx);
            return $"Cleared speaker notes for {targets.Length} slides. Output: {outputPath}";
        });
    }
}