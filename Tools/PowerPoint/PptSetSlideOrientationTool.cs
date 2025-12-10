using System.Text.Json.Nodes;
using Aspose.Slides;
using Aspose.Slides.Export;

namespace AsposeMcpServer.Tools;

public class PptSetSlideOrientationTool : IAsposeTool
{
    public string Description => "Set slide orientation (portrait or landscape)";

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
            orientation = new
            {
                type = "string",
                description = "Orientation: 'Portrait' or 'Landscape'",
                @enum = new[] { "Portrait", "Landscape" }
            }
        },
        required = new[] { "path", "orientation" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var orientation = arguments?["orientation"]?.GetValue<string>() ?? throw new ArgumentException("orientation is required");

        using var presentation = new Presentation(path);
        
        if (orientation.ToLower() == "portrait")
        {
            presentation.SlideSize.SetSize(SlideSizeType.A4Paper, SlideSizeScaleType.EnsureFit);
        }
        else
        {
            presentation.SlideSize.SetSize(SlideSizeType.OnScreen16x10, SlideSizeScaleType.EnsureFit);
        }

        presentation.Save(path, SaveFormat.Pptx);

        return await Task.FromResult($"Slide orientation set to {orientation}: {path}");
    }
}

