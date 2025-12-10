using System.Text.Json.Nodes;
using Aspose.Slides;
using Aspose.Slides.Export;

namespace AsposeMcpServer.Tools;

public class PptApplyThemeTool : IAsposeTool
{
    public string Description => "Apply a theme to a PowerPoint presentation";

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
            themePath = new
            {
                type = "string",
                description = "Theme template file path"
            }
        },
        required = new[] { "path", "themePath" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var themePath = arguments?["themePath"]?.GetValue<string>() ?? throw new ArgumentException("themePath is required");

        using var presentation = new Presentation(path);
        using var themePresentation = new Presentation(themePath);

        // Copy theme from the first slide of theme presentation
        presentation.Slides[0].LayoutSlide = themePresentation.Slides[0].LayoutSlide;

        presentation.Save(path, SaveFormat.Pptx);

        return await Task.FromResult($"Theme applied to presentation: {path}");
    }
}

