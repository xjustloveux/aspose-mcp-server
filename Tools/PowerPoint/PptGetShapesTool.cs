using System.Text.Json.Nodes;
using System.Text;
using Aspose.Slides;
using Aspose.Slides.Charts;
using Aspose.Slides.SmartArt;

namespace AsposeMcpServer.Tools;

public class PptGetShapesTool : IAsposeTool
{
    public string Description => "List shapes on a slide with type, text, position, and size";

    public object InputSchema => new
    {
        type = "object",
        properties = new
        {
            path = new { type = "string", description = "Presentation file path" },
            slideIndex = new { type = "number", description = "Slide index (0-based)" }
        },
        required = new[] { "path", "slideIndex" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var slideIndex = arguments?["slideIndex"]?.GetValue<int>() ?? throw new ArgumentException("slideIndex is required");

        using var presentation = new Presentation(path);
        if (slideIndex < 0 || slideIndex >= presentation.Slides.Count)
        {
            throw new ArgumentException($"slideIndex must be between 0 and {presentation.Slides.Count - 1}");
        }

        var slide = presentation.Slides[slideIndex];
        var sb = new StringBuilder();
        sb.AppendLine($"Slide {slideIndex} shapes: {slide.Shapes.Count}");

        for (int i = 0; i < slide.Shapes.Count; i++)
        {
            var s = slide.Shapes[i];
            var kind = s switch
            {
                IAutoShape => "AutoShape",
                PictureFrame => "Picture",
                ITable => "Table",
                IChart => "Chart",
                IGroupShape => "Group",
                ISmartArt => "SmartArt",
                IAudioFrame => "Audio",
                IVideoFrame => "Video",
                _ => s.GetType().Name
            };

            var text = (s as IAutoShape)?.TextFrame?.Text;
            sb.AppendLine($"[{i}] {kind} pos=({s.X},{s.Y}) size=({s.Width},{s.Height}) text={(string.IsNullOrWhiteSpace(text) ? "(none)" : text)}");
        }

        return await Task.FromResult(sb.ToString());
    }
}

