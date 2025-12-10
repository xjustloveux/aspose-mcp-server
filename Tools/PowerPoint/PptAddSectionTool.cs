using System.Text.Json.Nodes;
using Aspose.Slides;
using Aspose.Slides.Export;

namespace AsposeMcpServer.Tools;

public class PptAddSectionTool : IAsposeTool
{
    public string Description => "Add a new section starting at a slide index";

    public object InputSchema => new
    {
        type = "object",
        properties = new
        {
            path = new { type = "string", description = "Presentation file path" },
            name = new { type = "string", description = "Section name" },
            slideIndex = new { type = "number", description = "Start slide index for section (0-based)" }
        },
        required = new[] { "path", "name", "slideIndex" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var name = arguments?["name"]?.GetValue<string>() ?? throw new ArgumentException("name is required");
        var slideIndex = arguments?["slideIndex"]?.GetValue<int>() ?? throw new ArgumentException("slideIndex is required");

        using var presentation = new Presentation(path);
        if (slideIndex < 0 || slideIndex >= presentation.Slides.Count)
        {
            throw new ArgumentException($"slideIndex must be between 0 and {presentation.Slides.Count - 1}");
        }

        presentation.Sections.AddSection(name, presentation.Slides[slideIndex]);
        presentation.Save(path, SaveFormat.Pptx);
        return await Task.FromResult($"已新增章節 '{name}' 起始於投影片 {slideIndex}");
    }
}

