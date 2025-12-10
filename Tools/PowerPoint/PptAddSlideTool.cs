using System.Text.Json.Nodes;
using Aspose.Slides;
using Aspose.Slides.Export;

namespace AsposeMcpServer.Tools;

public class PptAddSlideTool : IAsposeTool
{
    public string Description => "Add a new slide to a PowerPoint presentation";

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
            layoutType = new
            {
                type = "string",
                description = "Slide layout type (Blank, Title, TitleOnly, etc., optional)"
            }
        },
        required = new[] { "path" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var layoutType = arguments?["layoutType"]?.GetValue<string>() ?? "Blank";

        using var presentation = new Presentation(path);
        
        if (presentation.LayoutSlides.Count == 0)
        {
            throw new InvalidOperationException("Presentation has no layout slides");
        }
        
        var layoutSlide = presentation.LayoutSlides[0];
        var slide = presentation.Slides.AddEmptySlide(layoutSlide);

        presentation.Save(path, SaveFormat.Pptx);

        return await Task.FromResult($"Slide added to presentation: {path} (Total: {presentation.Slides.Count})");
    }
}

