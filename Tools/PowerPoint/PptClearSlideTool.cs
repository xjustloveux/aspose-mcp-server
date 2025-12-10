using System.Text.Json.Nodes;
using Aspose.Slides;
using Aspose.Slides.Export;

namespace AsposeMcpServer.Tools;

public class PptClearSlideTool : IAsposeTool
{
    public string Description => "Clear all shapes from a slide";

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
            clearNotes = new
            {
                type = "boolean",
                description = "Also clear notes (optional, default: false)"
            }
        },
        required = new[] { "path", "slideIndex" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var slideIndex = arguments?["slideIndex"]?.GetValue<int>() ?? throw new ArgumentException("slideIndex is required");
        var clearNotes = arguments?["clearNotes"]?.GetValue<bool?>() ?? false;

        using var presentation = new Presentation(path);
        if (slideIndex < 0 || slideIndex >= presentation.Slides.Count)
        {
            throw new ArgumentException($"slideIndex must be between 0 and {presentation.Slides.Count - 1}");
        }

        var slide = presentation.Slides[slideIndex];
        var shapeCount = slide.Shapes.Count;

        // Remove all shapes
        while (slide.Shapes.Count > 0)
        {
            slide.Shapes.RemoveAt(slide.Shapes.Count - 1);
        }

        if (clearNotes)
        {
            var notesSlide = slide.NotesSlideManager.NotesSlide;
            if (notesSlide != null)
            {
                notesSlide.NotesTextFrame.Text = "";
            }
        }

        presentation.Save(path, SaveFormat.Pptx);

        return await Task.FromResult($"Cleared {shapeCount} shapes from slide {slideIndex}: {path}");
    }
}

