using System.Text.Json.Nodes;
using Aspose.Slides;
using Aspose.Slides.Export;
using System.Linq;

namespace AsposeMcpServer.Tools;

public class PptApplyLayoutRangeTool : IAsposeTool
{
    public string Description => "Apply a slide layout to multiple slides";

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
                description = "Slide indices to apply (0-based)"
            },
            layout = new
            {
                type = "string",
                description = "Layout type (Title, TitleOnly, Blank, TwoColumn, SectionHeader)"
            }
        },
        required = new[] { "path", "slideIndices", "layout" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var layoutStr = arguments?["layout"]?.GetValue<string>() ?? throw new ArgumentException("layout is required");
        var slideIndices = arguments?["slideIndices"]?.AsArray()?.Select(x => x?.GetValue<int>() ?? -1).ToArray()
                           ?? throw new ArgumentException("slideIndices is required");

        using var presentation = new Presentation(path);

        foreach (var idx in slideIndices)
        {
            if (idx < 0 || idx >= presentation.Slides.Count)
            {
                throw new ArgumentException($"slide index {idx} out of range");
            }
        }

        var layoutType = layoutStr.ToLower() switch
        {
            "title" => SlideLayoutType.Title,
            "titleonly" => SlideLayoutType.TitleOnly,
            "blank" => SlideLayoutType.Blank,
            "twocolumn" => SlideLayoutType.TwoColumnText,
            "sectionheader" => SlideLayoutType.SectionHeader,
            _ => SlideLayoutType.Custom
        };

        var layout = presentation.LayoutSlides.FirstOrDefault(ls => ls.LayoutType == layoutType) ?? presentation.LayoutSlides[0];

        foreach (var idx in slideIndices)
        {
            presentation.Slides[idx].LayoutSlide = layout;
        }

        presentation.Save(path, SaveFormat.Pptx);
        return await Task.FromResult($"已套用版面 {layoutStr} 到 {slideIndices.Length} 張投影片");
    }
}

