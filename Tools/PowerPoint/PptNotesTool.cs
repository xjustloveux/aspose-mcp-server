using System.Text.Json;
using System.Text.Json.Nodes;
using Aspose.Slides;
using Aspose.Slides.Export;
using AsposeMcpServer.Core;

namespace AsposeMcpServer.Tools.PowerPoint;

/// <summary>
///     Unified tool for managing PowerPoint notes.
///     Supports: add, edit, get, clear, set_header_footer
///     Note: Notes pages have separate header and footer fields (unlike slides which only have footer).
/// </summary>
public class PptNotesTool : IAsposeTool
{
    private static readonly JsonSerializerOptions JsonOptions = new() { WriteIndented = true };

    public string Description => @"Manage PowerPoint notes. Supports 4 operations: set, get, clear, set_header_footer.

Note: Notes pages have separate header and footer fields (unlike slides which only have footer).
Warning: 'set' operation will REPLACE existing notes content (format will be reset).
Warning: If outputPath is not provided, the original file will be overwritten.

Usage examples:
- Set notes: ppt_notes(operation='set', path='presentation.pptx', slideIndex=0, notes='Speaker notes')
- Get notes: ppt_notes(operation='get', path='presentation.pptx', slideIndex=0)
- Clear notes: ppt_notes(operation='clear', path='presentation.pptx', slideIndices=[0,1,2])
- Set header/footer: ppt_notes(operation='set_header_footer', path='presentation.pptx', headerText='Header', footerText='Footer')";

    public object InputSchema => new
    {
        type = "object",
        properties = new
        {
            operation = new
            {
                type = "string",
                description = @"Operation to perform.
- 'set': Set/replace notes on slide (required: path, slideIndex, notes)
- 'get': Get slide notes as JSON (required: path; optional: slideIndex)
- 'clear': Clear notes from slides (required: path; optional: slideIndices)
- 'set_header_footer': Set header/footer for notes master (required: path)",
                @enum = new[] { "set", "get", "clear", "set_header_footer" }
            },
            path = new
            {
                type = "string",
                description = "Presentation file path (required for all operations)"
            },
            slideIndex = new
            {
                type = "number",
                description = "Slide index (0-based, required for add/edit, optional for get)"
            },
            notes = new
            {
                type = "string",
                description = "Notes text content (required for add/edit). Will replace existing notes."
            },
            slideIndices = new
            {
                type = "array",
                items = new { type = "number" },
                description = "Slide indices array (optional for clear, if not provided affects all slides)"
            },
            headerText = new
            {
                type = "string",
                description = "Header text for notes pages (optional for set_header_footer)"
            },
            footerText = new
            {
                type = "string",
                description = "Footer text for notes pages (optional for set_header_footer)"
            },
            dateText = new
            {
                type = "string",
                description = "Date/time text for notes pages (optional for set_header_footer)"
            },
            showPageNumber = new
            {
                type = "boolean",
                description = "Show page number on notes pages (optional for set_header_footer, default: true)"
            },
            outputPath = new
            {
                type = "string",
                description = "Output file path (optional, defaults to input path - will OVERWRITE original file)"
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
            "set" => await SetNotesAsync(path, outputPath, arguments),
            "get" => await GetNotesAsync(path, arguments),
            "clear" => await ClearNotesAsync(path, outputPath, arguments),
            "set_header_footer" => await SetNotesHeaderFooterAsync(path, outputPath, arguments),
            _ => throw new ArgumentException($"Unknown operation: {operation}")
        };
    }

    /// <summary>
    ///     Sets (replaces) notes on a slide.
    /// </summary>
    /// <param name="path">PowerPoint file path.</param>
    /// <param name="outputPath">Output file path.</param>
    /// <param name="arguments">JSON arguments containing slideIndex, notes.</param>
    /// <returns>Success message.</returns>
    /// <exception cref="ArgumentException">Thrown when slideIndex is out of range.</exception>
    private Task<string> SetNotesAsync(string path, string outputPath, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var slideIndex = ArgumentHelper.GetInt(arguments, "slideIndex");
            var notes = ArgumentHelper.GetString(arguments, "notes");

            using var presentation = new Presentation(path);
            PowerPointHelper.ValidateCollectionIndex(slideIndex, presentation.Slides.Count, "slide");

            var slide = presentation.Slides[slideIndex];
            var notesSlide = GetOrAddNotesSlide(slide);
            notesSlide.NotesTextFrame.Text = notes;

            presentation.Save(outputPath, SaveFormat.Pptx);
            return $"Notes set for slide {slideIndex}. Output: {outputPath}";
        });
    }

    /// <summary>
    ///     Gets notes from slides.
    /// </summary>
    /// <param name="path">PowerPoint file path.</param>
    /// <param name="arguments">JSON arguments containing optional slideIndex.</param>
    /// <returns>JSON string with notes information.</returns>
    /// <exception cref="ArgumentException">Thrown when slideIndex is out of range.</exception>
    private Task<string> GetNotesAsync(string path, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var slideIndex = ArgumentHelper.GetIntNullable(arguments, "slideIndex");

            using var presentation = new Presentation(path);

            if (slideIndex.HasValue)
            {
                PowerPointHelper.ValidateCollectionIndex(slideIndex.Value, presentation.Slides.Count, "slide");

                var notesSlide = presentation.Slides[slideIndex.Value].NotesSlideManager.NotesSlide;
                var notesText = notesSlide?.NotesTextFrame?.Text;

                var result = new
                {
                    slideIndex = slideIndex.Value,
                    hasNotes = !string.IsNullOrWhiteSpace(notesText),
                    notes = notesText
                };

                return JsonSerializer.Serialize(result, JsonOptions);
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

                return JsonSerializer.Serialize(result, JsonOptions);
            }
        });
    }

    /// <summary>
    ///     Clears notes from slides. Only clears if notes slide exists.
    /// </summary>
    /// <param name="path">PowerPoint file path.</param>
    /// <param name="outputPath">Output file path.</param>
    /// <param name="arguments">JSON arguments containing optional slideIndices.</param>
    /// <returns>Success message with cleared count.</returns>
    /// <exception cref="ArgumentException">Thrown when any slideIndex is out of range.</exception>
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

            ValidateSlideIndices(targets, presentation.Slides.Count);

            var clearedCount = 0;
            foreach (var idx in targets)
            {
                var slide = presentation.Slides[idx];
                var notesSlide = slide.NotesSlideManager.NotesSlide;
                if (notesSlide?.NotesTextFrame != null)
                {
                    notesSlide.NotesTextFrame.Text = string.Empty;
                    clearedCount++;
                }
            }

            presentation.Save(outputPath, SaveFormat.Pptx);
            return
                $"Cleared speaker notes for {clearedCount} slides (of {targets.Length} targeted). Output: {outputPath}";
        });
    }

    /// <summary>
    ///     Sets header and footer for notes master.
    /// </summary>
    /// <param name="path">PowerPoint file path.</param>
    /// <param name="outputPath">Output file path.</param>
    /// <param name="arguments">JSON arguments containing headerText, footerText, dateText, showPageNumber.</param>
    /// <returns>Success message.</returns>
    /// <exception cref="InvalidOperationException">Thrown when presentation is empty or notes master creation fails.</exception>
    private Task<string> SetNotesHeaderFooterAsync(string path, string outputPath, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var headerText = ArgumentHelper.GetStringNullable(arguments, "headerText");
            var footerText = ArgumentHelper.GetStringNullable(arguments, "footerText");
            var dateText = ArgumentHelper.GetStringNullable(arguments, "dateText");
            var showPageNumber = ArgumentHelper.GetBool(arguments, "showPageNumber", true);

            using var presentation = new Presentation(path);

            if (presentation.Slides.Count == 0)
                throw new InvalidOperationException(
                    "Presentation has no slides. Cannot set notes header/footer on empty presentation.");

            if (presentation.MasterNotesSlideManager.MasterNotesSlide == null)
                presentation.Slides[0].NotesSlideManager.AddNotesSlide();

            var notesMaster = presentation.MasterNotesSlideManager.MasterNotesSlide;
            if (notesMaster == null)
                throw new InvalidOperationException("Failed to create notes master slide.");

            var manager = notesMaster.HeaderFooterManager;

            if (!string.IsNullOrEmpty(headerText))
            {
                manager.SetHeaderText(headerText);
                manager.SetHeaderVisibility(true);
            }

            if (!string.IsNullOrEmpty(footerText))
            {
                manager.SetFooterText(footerText);
                manager.SetFooterVisibility(true);
            }

            if (!string.IsNullOrEmpty(dateText))
            {
                manager.SetDateTimeText(dateText);
                manager.SetDateTimeVisibility(true);
            }

            manager.SetSlideNumberVisibility(showPageNumber);

            presentation.Save(outputPath, SaveFormat.Pptx);

            var settings = new List<string>();
            if (!string.IsNullOrEmpty(headerText)) settings.Add("header");
            if (!string.IsNullOrEmpty(footerText)) settings.Add("footer");
            if (!string.IsNullOrEmpty(dateText)) settings.Add("date");
            settings.Add(showPageNumber ? "page number shown" : "page number hidden");

            return $"Notes master header/footer updated ({string.Join(", ", settings)}). Output: {outputPath}";
        });
    }

    #region Helper Methods

    /// <summary>
    ///     Gets or creates a notes slide for the given slide.
    /// </summary>
    /// <param name="slide">The slide to get notes for.</param>
    /// <returns>The notes slide.</returns>
    private static INotesSlide GetOrAddNotesSlide(ISlide slide)
    {
        return slide.NotesSlideManager.NotesSlide ?? slide.NotesSlideManager.AddNotesSlide();
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

    #endregion
}