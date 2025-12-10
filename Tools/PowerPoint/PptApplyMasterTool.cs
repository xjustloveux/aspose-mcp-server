using System.Text.Json.Nodes;
using Aspose.Slides;
using Aspose.Slides.Export;

namespace AsposeMcpServer.Tools;

public class PptApplyMasterTool : IAsposeTool
{
    public string Description => "Apply a master (and its layout) to one or multiple slides";

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
                description = "Slide indices to apply (0-based, default all slides)"
            },
            masterIndex = new { type = "number", description = "Master index (0-based, default 0)" },
            layoutIndex = new { type = "number", description = "Layout index under master (0-based, default 0)" }
        },
        required = new[] { "path" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var slideIndices = arguments?["slideIndices"]?.AsArray()?.Select(x => x?.GetValue<int>() ?? -1).ToArray();
        var masterIndex = arguments?["masterIndex"]?.GetValue<int?>() ?? 0;
        var layoutIndex = arguments?["layoutIndex"]?.GetValue<int?>() ?? 0;

        using var presentation = new Presentation(path);

        if (masterIndex < 0 || masterIndex >= presentation.Masters.Count)
        {
            throw new ArgumentException($"masterIndex must be between 0 and {presentation.Masters.Count - 1}");
        }

        var master = presentation.Masters[masterIndex];
        if (layoutIndex < 0 || layoutIndex >= master.LayoutSlides.Count)
        {
            throw new ArgumentException($"layoutIndex must be between 0 and {master.LayoutSlides.Count - 1}");
        }

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

        var layout = master.LayoutSlides[layoutIndex];
        foreach (var idx in targets)
        {
            presentation.Slides[idx].LayoutSlide = layout;
        }

        presentation.Save(path, SaveFormat.Pptx);
        return await Task.FromResult($"已套用母片 {masterIndex} / 版面 {layoutIndex} 至 {targets.Length} 張投影片");
    }
}

