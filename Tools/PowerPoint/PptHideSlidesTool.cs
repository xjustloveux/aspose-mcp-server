using System.Text.Json.Nodes;
using Aspose.Slides;
using Aspose.Slides.Export;
using System.Linq;

namespace AsposeMcpServer.Tools;

public class PptHideSlidesTool : IAsposeTool
{
    public string Description => "Hide or show selected slides";

    public object InputSchema => new
    {
        type = "object",
        properties = new
        {
            path = new { type = "string", description = "Presentation file path" },
            slideIndices = new
            {
                type = "array",
                items = new { type = "number" },
                description = "Slide indices to update (0-based; default all)"
            },
            hidden = new { type = "boolean", description = "Hide slides (true) or show (false)" }
        },
        required = new[] { "path", "hidden" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var hidden = arguments?["hidden"]?.GetValue<bool?>() ?? false;
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
            presentation.Slides[idx].Hidden = hidden;
        }

        presentation.Save(path, SaveFormat.Pptx);
        return await Task.FromResult($"已設定 {targets.Length} 張投影片 Hidden={hidden}");
    }
}

