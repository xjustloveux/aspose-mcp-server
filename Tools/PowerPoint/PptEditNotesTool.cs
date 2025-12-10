using System.Text.Json.Nodes;
using Aspose.Slides;
using Aspose.Slides.Export;

namespace AsposeMcpServer.Tools;

public class PptEditNotesTool : IAsposeTool
{
    public string Description => "Edit slide notes text";

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
            notes = new
            {
                type = "string",
                description = "Notes text"
            }
        },
        required = new[] { "path", "slideIndex", "notes" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var slideIndex = arguments?["slideIndex"]?.GetValue<int>() ?? throw new ArgumentException("slideIndex is required");
        var notes = arguments?["notes"]?.GetValue<string>() ?? throw new ArgumentException("notes is required");

        using var presentation = new Presentation(path);
        if (slideIndex < 0 || slideIndex >= presentation.Slides.Count)
        {
            throw new ArgumentException($"slideIndex must be between 0 and {presentation.Slides.Count - 1}");
        }

        var slide = presentation.Slides[slideIndex];
        var notesSlide = slide.NotesSlideManager.NotesSlide;
        if (notesSlide == null)
        {
            notesSlide = slide.NotesSlideManager.AddNotesSlide();
        }
        notesSlide.NotesTextFrame.Text = notes;

        presentation.Save(path, SaveFormat.Pptx);

        return await Task.FromResult($"Notes updated for slide {slideIndex}: {path}");
    }
}

