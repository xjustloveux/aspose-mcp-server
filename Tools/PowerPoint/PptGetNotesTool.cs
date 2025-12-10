using System.Text.Json.Nodes;
using System.Text;
using Aspose.Slides;

namespace AsposeMcpServer.Tools;

public class PptGetNotesTool : IAsposeTool
{
    public string Description => "Get speaker notes from a PowerPoint slide";

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
                description = "Slide index (0-based, optional, if not provided returns all notes)"
            }
        },
        required = new[] { "path" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var slideIndex = arguments?["slideIndex"]?.GetValue<int?>();

        using var presentation = new Presentation(path);
        var sb = new StringBuilder();

        if (slideIndex.HasValue)
        {
            if (slideIndex.Value < 0 || slideIndex.Value >= presentation.Slides.Count)
            {
                throw new ArgumentException($"slideIndex must be between 0 and {presentation.Slides.Count - 1}");
            }

            var slide = presentation.Slides[slideIndex.Value];
            var notesSlide = presentation.Slides[slideIndex.Value].NotesSlideManager.NotesSlide;
            
            if (notesSlide != null && notesSlide.NotesTextFrame != null)
            {
                sb.AppendLine($"Slide {slideIndex.Value} Notes:");
                sb.AppendLine(notesSlide.NotesTextFrame.Text);
            }
            else
            {
                sb.AppendLine($"Slide {slideIndex.Value} has no notes.");
            }
        }
        else
        {
            sb.AppendLine("All Speaker Notes:");
            for (int i = 0; i < presentation.Slides.Count; i++)
            {
                var notesSlide = presentation.Slides[i].NotesSlideManager.NotesSlide;
                if (notesSlide != null && notesSlide.NotesTextFrame != null && !string.IsNullOrWhiteSpace(notesSlide.NotesTextFrame.Text))
                {
                    sb.AppendLine($"\n--- Slide {i} ---");
                    sb.AppendLine(notesSlide.NotesTextFrame.Text);
                }
            }
        }

        return await Task.FromResult(sb.ToString());
    }
}

