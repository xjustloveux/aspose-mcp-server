using System.Text.Json.Nodes;
using Aspose.Slides;
using Aspose.Slides.Export;
using System.Linq;

namespace AsposeMcpServer.Tools;

public class PptSetLayoutTool : IAsposeTool
{
    public string Description => "Set a slide layout (Title, TitleOnly, Blank, TwoColumn, etc.)";

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
                description = "Layout type (Title, TitleOnly, Blank, TwoColumn, SectionHeader, Comparison, PictureWithCaption, Custom)"
            }
        },
        required = new[] { "path", "slideIndex", "layout" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var slideIndex = arguments?["slideIndex"]?.GetValue<int>() ?? throw new ArgumentException("slideIndex is required");
        var layoutStr = arguments?["layout"]?.GetValue<string>() ?? throw new ArgumentException("layout is required");

        using var presentation = new Presentation(path);
        if (slideIndex < 0 || slideIndex >= presentation.Slides.Count)
        {
            throw new ArgumentException($"slideIndex must be between 0 and {presentation.Slides.Count - 1}");
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
        presentation.Slides[slideIndex].LayoutSlide = layout;

        presentation.Save(path, SaveFormat.Pptx);
        return await Task.FromResult($"已設定投影片 {slideIndex} 版面：{layoutStr}");
    }
}

