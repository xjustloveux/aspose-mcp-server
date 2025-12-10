using System.Text.Json.Nodes;
using Aspose.Slides;
using Aspose.Slides.Export;

namespace AsposeMcpServer.Tools;

public class PptEditSlideTool : IAsposeTool
{
    public string Description => "Edit slide properties (hidden status, notes, etc.)";

    public object InputSchema => new
    {
        type = "object",
        properties = new
        {
            path = new
            {
                type = "string",
                description = "Presentation file path"
            },
            slideIndex = new
            {
                type = "number",
                description = "Slide index (0-based)"
            },
            hidden = new
            {
                type = "boolean",
                description = "Hide/show slide (optional)"
            },
            notes = new
            {
                type = "string",
                description = "Slide notes text (optional)"
            }
        },
        required = new[] { "path", "slideIndex" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var slideIndex = arguments?["slideIndex"]?.GetValue<int>() ?? throw new ArgumentException("slideIndex is required");
        var hidden = arguments?["hidden"]?.GetValue<bool?>();
        var notes = arguments?["notes"]?.GetValue<string>();

        using var presentation = new Presentation(path);
        if (slideIndex < 0 || slideIndex >= presentation.Slides.Count)
        {
            throw new ArgumentException($"slideIndex must be between 0 and {presentation.Slides.Count - 1}");
        }

        var slide = presentation.Slides[slideIndex];
        var changes = new List<string>();

        if (hidden.HasValue)
        {
            slide.Hidden = hidden.Value;
            changes.Add($"Hidden: {hidden.Value}");
        }

        if (notes != null)
        {
            var notesSlide = slide.NotesSlideManager.NotesSlide;
            if (notesSlide == null)
            {
                notesSlide = slide.NotesSlideManager.AddNotesSlide();
            }
            notesSlide.NotesTextFrame.Text = notes;
            changes.Add("Notes updated");
        }

        presentation.Save(path, SaveFormat.Pptx);

        return await Task.FromResult($"Slide {slideIndex} edited: {string.Join(", ", changes)} - {path}");
    }
}

