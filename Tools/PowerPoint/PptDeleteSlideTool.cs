using System.Text.Json.Nodes;
using Aspose.Slides;
using Aspose.Slides.Export;

namespace AsposeMcpServer.Tools;

public class PptDeleteSlideTool : IAsposeTool
{
    public string Description => "Delete a slide from a PowerPoint presentation";

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
                description = "Slide index to delete (0-based)"
            }
        },
        required = new[] { "path", "slideIndex" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var slideIndex = arguments?["slideIndex"]?.GetValue<int>() ?? throw new ArgumentException("slideIndex is required");

        using var presentation = new Presentation(path);
        if (slideIndex < 0 || slideIndex >= presentation.Slides.Count)
        {
            throw new ArgumentException($"slideIndex must be between 0 and {presentation.Slides.Count - 1}");
        }

        presentation.Slides.RemoveAt(slideIndex);
        presentation.Save(path, SaveFormat.Pptx);

        return await Task.FromResult($"已刪除投影片 {slideIndex}，剩餘 {presentation.Slides.Count} 張: {path}");
    }
}

