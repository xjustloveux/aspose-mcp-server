using System.Text.Json.Nodes;
using System.Text;
using Aspose.Slides;

namespace AsposeMcpServer.Tools;

public class PptGetSlidesInfoTool : IAsposeTool
{
    public string Description => "Get slide-level info: titles, shape count, notes presence";

    public object InputSchema => new
    {
        type = "object",
        properties = new
        {
            path = new
            {
                type = "string",
                description = "Presentation file path"
            }
        },
        required = new[] { "path" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");

        using var presentation = new Presentation(path);
        var sb = new StringBuilder();

        sb.AppendLine($"總投影片數: {presentation.Slides.Count}");

        for (int i = 0; i < presentation.Slides.Count; i++)
        {
            var slide = presentation.Slides[i];
            var title = slide.Shapes.FirstOrDefault(s => s.Placeholder?.Type == PlaceholderType.Title) as IAutoShape;
            var titleText = title?.TextFrame?.Text ?? "(無標題)";
            var notes = slide.NotesSlideManager.NotesSlide?.NotesTextFrame?.Text;

            sb.AppendLine($"\n--- 投影片 {i} ---");
            sb.AppendLine($"標題: {titleText}");
            sb.AppendLine($"形狀數: {slide.Shapes.Count}");
            sb.AppendLine($"是否有講者備註: {(string.IsNullOrWhiteSpace(notes) ? "否" : "是")}");
        }

        return await Task.FromResult(sb.ToString());
    }
}

