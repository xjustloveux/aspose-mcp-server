using System.Text.Json.Nodes;
using System.Text;
using Aspose.Slides;
using Aspose.Slides.Export;

namespace AsposeMcpServer.Tools;

public class PptGetMasterSlidesTool : IAsposeTool
{
    public string Description => "Get all master slides and their layouts";

    public object InputSchema => new
    {
        type = "object",
        properties = new
        {
            path = new
            {
                type = "string",
                description = "Presentation file path"
            }
        },
        required = new[] { "path" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");

        using var presentation = new Presentation(path);
        var sb = new StringBuilder();

        sb.AppendLine($"=== Master Slides ===");
        sb.AppendLine($"Total: {presentation.Masters.Count}");

        for (int i = 0; i < presentation.Masters.Count; i++)
        {
            var master = presentation.Masters[i];
            sb.AppendLine($"\nMaster {i}:");
            sb.AppendLine($"  Name: {master.Name ?? "(unnamed)"}");
            sb.AppendLine($"  Layouts: {master.LayoutSlides.Count}");
            for (int j = 0; j < master.LayoutSlides.Count; j++)
            {
                var layout = master.LayoutSlides[j];
                sb.AppendLine($"    [{j}] {layout.Name ?? "(unnamed)"}");
            }
        }

        return await Task.FromResult(sb.ToString());
    }
}

