using System.ComponentModel;
using System.Text.Json;
using Aspose.Slides;
using AsposeMcpServer.Core.Helpers;
using AsposeMcpServer.Core.Session;
using ModelContextProtocol.Server;

namespace AsposeMcpServer.Tools.PowerPoint;

/// <summary>
///     Unified tool for managing PowerPoint notes.
///     Supports: set, get, clear, set_header_footer
///     Note: Notes pages have separate header and footer fields (unlike slides which only have footer).
/// </summary>
[McpServerToolType]
public class PptNotesTool
{
    /// <summary>
    ///     JSON serializer options for consistent output formatting.
    /// </summary>
    private static readonly JsonSerializerOptions JsonOptions = new() { WriteIndented = true };

    /// <summary>
    ///     Session manager for document session handling.
    /// </summary>
    private readonly DocumentSessionManager? _sessionManager;

    /// <summary>
    ///     Initializes a new instance of the <see cref="PptNotesTool" /> class.
    /// </summary>
    /// <param name="sessionManager">Optional session manager for in-memory editing.</param>
    public PptNotesTool(DocumentSessionManager? sessionManager = null)
    {
        _sessionManager = sessionManager;
    }

    [McpServerTool(Name = "ppt_notes")]
    [Description(@"Manage PowerPoint notes. Supports 4 operations: set, get, clear, set_header_footer.

Note: Notes pages have separate header and footer fields (unlike slides which only have footer).
Warning: 'set' operation will REPLACE existing notes content (format will be reset).
Warning: If outputPath is not provided, the original file will be overwritten.

Usage examples:
- Set notes: ppt_notes(operation='set', path='presentation.pptx', slideIndex=0, notes='Speaker notes')
- Get notes: ppt_notes(operation='get', path='presentation.pptx', slideIndex=0)
- Clear notes: ppt_notes(operation='clear', path='presentation.pptx', slideIndices=[0,1,2])
- Set header/footer: ppt_notes(operation='set_header_footer', path='presentation.pptx', headerText='Header', footerText='Footer')")]
    public string Execute(
        [Description("Operation: set, get, clear, set_header_footer")]
        string operation,
        [Description("Presentation file path (required if no sessionId)")]
        string? path = null,
        [Description("Session ID for in-memory editing")]
        string? sessionId = null,
        [Description("Output file path (file mode only)")]
        string? outputPath = null,
        [Description("Slide index (0-based, required for set, optional for get)")]
        int? slideIndex = null,
        [Description("Notes text content (required for set). Will replace existing notes.")]
        string? notes = null,
        [Description("Slide indices array (optional for clear, if not provided affects all slides)")]
        int[]? slideIndices = null,
        [Description("Header text for notes pages (optional for set_header_footer)")]
        string? headerText = null,
        [Description("Footer text for notes pages (optional for set_header_footer)")]
        string? footerText = null,
        [Description("Date/time text for notes pages (optional for set_header_footer)")]
        string? dateText = null,
        [Description("Show page number on notes pages (optional for set_header_footer, default: true)")]
        bool showPageNumber = true)
    {
        using var ctx = DocumentContext<Presentation>.Create(_sessionManager, sessionId, path);

        return operation.ToLower() switch
        {
            "set" => SetNotes(ctx, outputPath, slideIndex, notes),
            "get" => GetNotes(ctx, slideIndex),
            "clear" => ClearNotes(ctx, outputPath, slideIndices),
            "set_header_footer" => SetNotesHeaderFooter(ctx, outputPath, headerText, footerText, dateText,
                showPageNumber),
            _ => throw new ArgumentException($"Unknown operation: {operation}")
        };
    }

    /// <summary>
    ///     Sets (replaces) notes on a slide.
    /// </summary>
    /// <param name="ctx">The document context.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="slideIndex">The slide index (0-based).</param>
    /// <param name="notes">The notes text content.</param>
    /// <returns>A message indicating the result of the operation.</returns>
    /// <exception cref="ArgumentException">Thrown when slideIndex or notes is not provided.</exception>
    private static string SetNotes(DocumentContext<Presentation> ctx, string? outputPath, int? slideIndex,
        string? notes)
    {
        if (!slideIndex.HasValue)
            throw new ArgumentException("slideIndex is required for set operation");
        if (string.IsNullOrEmpty(notes))
            throw new ArgumentException("notes is required for set operation");

        var presentation = ctx.Document;
        PowerPointHelper.ValidateCollectionIndex(slideIndex.Value, presentation.Slides.Count, "slide");

        var slide = presentation.Slides[slideIndex.Value];
        var notesSlide = GetOrAddNotesSlide(slide);
        notesSlide.NotesTextFrame.Text = notes;

        ctx.Save(outputPath);

        var result = $"Notes set for slide {slideIndex}.\n";
        result += ctx.GetOutputMessage(outputPath);
        return result;
    }

    /// <summary>
    ///     Gets notes from slides.
    /// </summary>
    /// <param name="ctx">The document context.</param>
    /// <param name="slideIndex">The slide index (0-based), or null to get all slides.</param>
    /// <returns>A JSON string containing notes information.</returns>
    private static string GetNotes(DocumentContext<Presentation> ctx, int? slideIndex)
    {
        var presentation = ctx.Document;

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
            List<object> notesList = [];
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
    }

    /// <summary>
    ///     Clears notes from slides. Only clears if notes slide exists.
    /// </summary>
    /// <param name="ctx">The document context.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="slideIndices">The array of slide indices to clear, or null for all slides.</param>
    /// <returns>A message indicating the result of the operation.</returns>
    private static string ClearNotes(DocumentContext<Presentation> ctx, string? outputPath, int[]? slideIndices)
    {
        var presentation = ctx.Document;
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

        ctx.Save(outputPath);

        var result = $"Cleared speaker notes for {clearedCount} slides (of {targets.Length} targeted).\n";
        result += ctx.GetOutputMessage(outputPath);
        return result;
    }

    /// <summary>
    ///     Sets header and footer for notes master.
    /// </summary>
    /// <param name="ctx">The document context.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="headerText">The header text for notes pages.</param>
    /// <param name="footerText">The footer text for notes pages.</param>
    /// <param name="dateText">The date/time text for notes pages.</param>
    /// <param name="showPageNumber">Whether to show page numbers.</param>
    /// <returns>A message indicating the result of the operation.</returns>
    /// <exception cref="InvalidOperationException">
    ///     Thrown when the presentation has no slides or notes master cannot be
    ///     created.
    /// </exception>
    private static string SetNotesHeaderFooter(DocumentContext<Presentation> ctx, string? outputPath,
        string? headerText, string? footerText, string? dateText, bool showPageNumber)
    {
        var presentation = ctx.Document;

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

        ctx.Save(outputPath);

        List<string> settings = [];
        if (!string.IsNullOrEmpty(headerText)) settings.Add("header");
        if (!string.IsNullOrEmpty(footerText)) settings.Add("footer");
        if (!string.IsNullOrEmpty(dateText)) settings.Add("date");
        settings.Add(showPageNumber ? "page number shown" : "page number hidden");

        var result = $"Notes master header/footer updated ({string.Join(", ", settings)}).\n";
        result += ctx.GetOutputMessage(outputPath);
        return result;
    }

    #region Helper Methods

    /// <summary>
    ///     Gets or creates a notes slide for the given slide.
    /// </summary>
    /// <param name="slide">The slide to get or create notes for.</param>
    /// <returns>The notes slide for the given slide.</returns>
    private static INotesSlide GetOrAddNotesSlide(ISlide slide)
    {
        return slide.NotesSlideManager.NotesSlide ?? slide.NotesSlideManager.AddNotesSlide();
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

    #endregion
}