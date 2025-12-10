using System.Text.Json.Nodes;
using Aspose.Slides;
using Aspose.Slides.Export;

namespace AsposeMcpServer.Tools;

public class PptRenameSectionTool : IAsposeTool
{
    public string Description => "Rename a section by index";

    public object InputSchema => new
    {
        type = "object",
        properties = new
        {
            path = new { type = "string", description = "Presentation file path" },
            sectionIndex = new { type = "number", description = "Section index (0-based)" },
            newName = new { type = "string", description = "New section name" }
        },
        required = new[] { "path", "sectionIndex", "newName" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var sectionIndex = arguments?["sectionIndex"]?.GetValue<int>() ?? throw new ArgumentException("sectionIndex is required");
        var newName = arguments?["newName"]?.GetValue<string>() ?? throw new ArgumentException("newName is required");

        using var presentation = new Presentation(path);
        if (sectionIndex < 0 || sectionIndex >= presentation.Sections.Count)
        {
            throw new ArgumentException($"sectionIndex must be between 0 and {presentation.Sections.Count - 1}");
        }

        presentation.Sections[sectionIndex].Name = newName;
        presentation.Save(path, SaveFormat.Pptx);
        return await Task.FromResult($"已重新命名章節 {sectionIndex} 為 '{newName}'");
    }
}

