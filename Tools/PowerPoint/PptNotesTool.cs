using System.Text.Json.Nodes;
using System.Text;
using System.Linq;
using Aspose.Slides;
using Aspose.Slides.Export;
using AsposeMcpServer.Core;

namespace AsposeMcpServer.Tools;

/// <summary>
/// Unified tool for managing PowerPoint notes (add, edit, get, clear)
/// Merges: PptAddNotesTool, PptEditNotesTool, PptGetNotesTool, PptClearNotesTool
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
            }
        },
        required = new[] { "operation", "path" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var operation = arguments?["operation"]?.GetValue<string>() ?? throw new ArgumentException("operation is required");
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");

        return operation.ToLower() switch
        {
            "add" => await AddNotesAsync(arguments, path),
            "edit" => await EditNotesAsync(arguments, path),
            "get" => await GetNotesAsync(arguments, path),
            "clear" => await ClearNotesAsync(arguments, path),
            _ => throw new ArgumentException($"Unknown operation: {operation}")
        };
    }

    private async Task<string> AddNotesAsync(JsonObject? arguments, string path)
    {
        var slideIndex = arguments?["slideIndex"]?.GetValue<int>() ?? throw new ArgumentException("slideIndex is required for add operation");
        var notes = arguments?["notes"]?.GetValue<string>() ?? throw new ArgumentException("notes is required for add operation");

        using var presentation = new Presentation(path);
        var slide = PowerPointHelper.GetSlide(presentation, slideIndex);
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

    private async Task<string> EditNotesAsync(JsonObject? arguments, string path)
    {
        var slideIndex = arguments?["slideIndex"]?.GetValue<int>() ?? throw new ArgumentException("slideIndex is required for edit operation");
        var notes = arguments?["notes"]?.GetValue<string>() ?? throw new ArgumentException("notes is required for edit operation");

        using var presentation = new Presentation(path);
        var slide = PowerPointHelper.GetSlide(presentation, slideIndex);
        var notesSlide = slide.NotesSlideManager.NotesSlide;
        if (notesSlide == null)
        {
            notesSlide = slide.NotesSlideManager.AddNotesSlide();
        }
        notesSlide.NotesTextFrame.Text = notes;

        presentation.Save(path, SaveFormat.Pptx);

        return await Task.FromResult($"Notes updated for slide {slideIndex}: {path}");
    }

    private async Task<string> GetNotesAsync(JsonObject? arguments, string path)
    {
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

    private async Task<string> ClearNotesAsync(JsonObject? arguments, string path)
    {
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

