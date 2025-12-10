using System.Text.Json.Nodes;
using System.Text;
using Aspose.Slides;

namespace AsposeMcpServer.Tools;

public class PptGetSectionsTool : IAsposeTool
{
    public string Description => "Get all sections with name and slide count";

    public object InputSchema => new
    {
        type = "object",
        properties = new
        {
            path = new { type = "string", description = "Presentation file path" }
        },
        required = new[] { "path" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");

        using var presentation = new Presentation(path);
        var sb = new StringBuilder();
        sb.AppendLine($"Sections: {presentation.Sections.Count}");
        for (int i = 0; i < presentation.Sections.Count; i++)
        {
            var sec = presentation.Sections[i];
            sb.AppendLine($"[{i}] {sec.Name}");
        }
        return await Task.FromResult(sb.ToString());
    }
}

