using System.Text.Json.Nodes;
using System.Text;
using Aspose.Slides;
using Aspose.Slides.Export;

namespace AsposeMcpServer.Tools;

public class PptGetLayoutsTool : IAsposeTool
{
    public string Description => "Get all layout slides in the presentation";

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
            masterIndex = new
            {
                type = "number",
                description = "Master index (0-based, optional, if not provided gets all layouts)"
            }
        },
        required = new[] { "path" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var masterIndex = arguments?["masterIndex"]?.GetValue<int?>();

        using var presentation = new Presentation(path);
        var sb = new StringBuilder();

        if (masterIndex.HasValue)
        {
            if (masterIndex.Value < 0 || masterIndex.Value >= presentation.Masters.Count)
            {
                throw new ArgumentException($"masterIndex must be between 0 and {presentation.Masters.Count - 1}");
            }
            var master = presentation.Masters[masterIndex.Value];
            sb.AppendLine($"=== Master {masterIndex.Value} Layouts ===");
            sb.AppendLine($"Total: {master.LayoutSlides.Count}");
            for (int i = 0; i < master.LayoutSlides.Count; i++)
            {
                var layout = master.LayoutSlides[i];
                sb.AppendLine($"  [{i}] {layout.Name ?? "(unnamed)"}");
            }
        }
        else
        {
            sb.AppendLine("=== All Layouts ===");
            for (int i = 0; i < presentation.Masters.Count; i++)
            {
                var master = presentation.Masters[i];
                sb.AppendLine($"\nMaster {i}: {master.LayoutSlides.Count} layout(s)");
                for (int j = 0; j < master.LayoutSlides.Count; j++)
                {
                    var layout = master.LayoutSlides[j];
                    sb.AppendLine($"  [{j}] {layout.Name ?? "(unnamed)"}");
                }
            }
        }

        return await Task.FromResult(sb.ToString());
    }
}

