using System.Text.Json.Nodes;
using System.Text;
using Aspose.Slides;

namespace AsposeMcpServer.Tools;

public class PptGetContentTool : IAsposeTool
{
    public string Description => "Get text content from a PowerPoint presentation";

    public object InputSchema => new
    {
        type = "object",
        properties = new
        {
            path = new
            {
                type = "string",
                description = "Input file path"
            }
        },
        required = new[] { "path" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");

        using var presentation = new Presentation(path);
        var sb = new StringBuilder();

        sb.AppendLine($"Total slides: {presentation.Slides.Count}");
        
        for (int i = 0; i < presentation.Slides.Count; i++)
        {
            var slide = presentation.Slides[i];
            sb.AppendLine($"\n--- Slide {i + 1} ---");
            
            foreach (var shape in slide.Shapes)
            {
                if (shape is IAutoShape autoShape && autoShape.TextFrame != null)
                {
                    sb.AppendLine(autoShape.TextFrame.Text);
                }
            }
        }

        return await Task.FromResult(sb.ToString());
    }
}

