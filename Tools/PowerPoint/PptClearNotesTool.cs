using System.Text.Json.Nodes;
using Aspose.Slides;
using Aspose.Slides.Export;

namespace AsposeMcpServer.Tools;

public class PptClearNotesTool : IAsposeTool
{
    public string Description => "Clear speaker notes on selected slides (or all)";

    public object InputSchema => new
    {
        type = "object",
        properties = new
        {
            path = new { type = "string", description = "Presentation file path" },
            slideIndices = new
            {
                type = "array",
                items = new { type = "number" },
                description = "Slide indices to clear (optional; default all)"
            }
        },
        required = new[] { "path" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
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
            var slide = presentation.Slides[idx];
            var notes = slide.NotesSlideManager.NotesSlide;
            if (notes != null && notes.NotesTextFrame != null)
            {
                notes.NotesTextFrame.Text = string.Empty;
            }
        }

        presentation.Save(path, SaveFormat.Pptx);
        return await Task.FromResult($"已清空 {targets.Length} 張投影片的講者備註");
    }
}

