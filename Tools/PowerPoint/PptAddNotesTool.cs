using System.Text.Json.Nodes;
using Aspose.Slides;
using Aspose.Slides.Export;

namespace AsposeMcpServer.Tools;

public class PptAddNotesTool : IAsposeTool
{
    public string Description => "Add or replace speaker notes for a slide";

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
                description = "Notes text content"
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
        var notesSlide = slide.NotesSlideManager.NotesSlide ?? slide.NotesSlideManager.AddNotesSlide();
        var textFrame = notesSlide.NotesTextFrame;
        if (textFrame == null)
        {
            throw new InvalidOperationException("無法取得 NotesTextFrame，可能是檔案受損或格式不支援");
        }
        textFrame.Paragraphs.Clear();
        var para = new Paragraph();
        para.Portions.Add(new Portion(notes));
        textFrame.Paragraphs.Add(para);

        presentation.Save(path, SaveFormat.Pptx);
        return await Task.FromResult($"已更新投影片 {slideIndex} 的講者備註: {path}");
    }
}

