using System.Text.Json.Nodes;
using Aspose.Slides;
using Aspose.Slides.Export;

namespace AsposeMcpServer.Tools;

public class PptDeleteSectionTool : IAsposeTool
{
    public string Description => "Delete a section by index";

    public object InputSchema => new
    {
        type = "object",
        properties = new
        {
            path = new { type = "string", description = "Presentation file path" },
            sectionIndex = new { type = "number", description = "Section index (0-based)" },
            keepSlides = new { type = "boolean", description = "Keep slides in presentation (default true)" }
        },
        required = new[] { "path", "sectionIndex" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var sectionIndex = arguments?["sectionIndex"]?.GetValue<int>() ?? throw new ArgumentException("sectionIndex is required");
        var keepSlides = arguments?["keepSlides"]?.GetValue<bool?>() ?? true;

        using var presentation = new Presentation(path);
        if (sectionIndex < 0 || sectionIndex >= presentation.Sections.Count)
        {
            throw new ArgumentException($"sectionIndex must be between 0 and {presentation.Sections.Count - 1}");
        }

        var section = presentation.Sections[sectionIndex];
        if (keepSlides)
        {
            presentation.Sections.RemoveSection(section);
        }
        else
        {
            presentation.Sections.RemoveSectionWithSlides(section);
        }
        presentation.Save(path, SaveFormat.Pptx);
        return await Task.FromResult($"已移除章節 {sectionIndex}，保留投影片: {keepSlides}");
    }
}

