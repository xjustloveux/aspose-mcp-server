using System.Text.Json.Nodes;
using Aspose.Slides;
using Aspose.Slides.Export;

namespace AsposeMcpServer.Tools;

public class PptDuplicateSlideTool : IAsposeTool
{
    public string Description => "Duplicate a slide to a specified position";

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
                description = "Slide index to duplicate (0-based)"
            },
            insertAt = new
            {
                type = "number",
                description = "Target index to insert clone (0-based, optional, default: append)"
            }
        },
        required = new[] { "path", "slideIndex" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var slideIndex = arguments?["slideIndex"]?.GetValue<int>() ?? throw new ArgumentException("slideIndex is required");
        var insertAt = arguments?["insertAt"]?.GetValue<int?>();

        using var presentation = new Presentation(path);
        var count = presentation.Slides.Count;

        if (slideIndex < 0 || slideIndex >= count)
        {
            throw new ArgumentException($"slideIndex must be between 0 and {count - 1}");
        }

        if (insertAt.HasValue)
        {
            if (insertAt.Value < 0 || insertAt.Value > count)
            {
                throw new ArgumentException($"insertAt must be between 0 and {count}");
            }

            presentation.Slides.InsertClone(insertAt.Value, presentation.Slides[slideIndex]);
        }
        else
        {
            presentation.Slides.AddClone(presentation.Slides[slideIndex]);
        }

        presentation.Save(path, SaveFormat.Pptx);
        return await Task.FromResult($"已複製投影片 {slideIndex}，總數 {presentation.Slides.Count} 張: {path}");
    }
}

