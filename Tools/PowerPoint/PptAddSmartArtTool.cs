using System.Text.Json.Nodes;
using Aspose.Slides;
using Aspose.Slides.SmartArt;
using Aspose.Slides.Export;

namespace AsposeMcpServer.Tools;

public class PptAddSmartArtTool : IAsposeTool
{
    public string Description => "Insert a SmartArt diagram (process, cycle, hierarchy, list, etc.)";

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
                description = "Slide index (0-based)"
            },
            layout = new
            {
                type = "string",
                description = "Layout: BasicProcess, ContinuousCycle, Hierarchy, BasicBlockList, etc."
            },
            x = new { type = "number", description = "X position (optional, default: 50)" },
            y = new { type = "number", description = "Y position (optional, default: 50)" },
            width = new { type = "number", description = "Width (optional, default: 400)" },
            height = new { type = "number", description = "Height (optional, default: 300)" }
        },
        required = new[] { "path", "slideIndex", "layout" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var slideIndex = arguments?["slideIndex"]?.GetValue<int>() ?? throw new ArgumentException("slideIndex is required");
        var layoutStr = arguments?["layout"]?.GetValue<string>() ?? throw new ArgumentException("layout is required");
        var x = arguments?["x"]?.GetValue<float?>() ?? 50;
        var y = arguments?["y"]?.GetValue<float?>() ?? 50;
        var width = arguments?["width"]?.GetValue<float?>() ?? 400;
        var height = arguments?["height"]?.GetValue<float?>() ?? 300;

        using var presentation = new Presentation(path);
        if (slideIndex < 0 || slideIndex >= presentation.Slides.Count)
        {
            throw new ArgumentException($"slideIndex must be between 0 and {presentation.Slides.Count - 1}");
        }

        var layout = layoutStr.ToLower() switch
        {
            "basicprocess" => SmartArtLayoutType.BasicProcess,
            "continuouscycle" => SmartArtLayoutType.ContinuousCycle,
            "hierarchy" => SmartArtLayoutType.Hierarchy,
            "basicblocklist" => SmartArtLayoutType.BasicBlockList,
            "basicpyramid" => SmartArtLayoutType.BasicPyramid,
            "stackedlist" => SmartArtLayoutType.StackedList,
            "horizontalmultilevelhierarchy" => SmartArtLayoutType.HorizontalMultiLevelHierarchy,
            _ => SmartArtLayoutType.BasicProcess
        };

        var slide = presentation.Slides[slideIndex];
        slide.Shapes.AddSmartArt(x, y, width, height, layout);

        presentation.Save(path, SaveFormat.Pptx);
        return await Task.FromResult($"已新增 SmartArt ({layout}) 至投影片 {slideIndex}");
    }
}

