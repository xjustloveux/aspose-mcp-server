using System.Text.Json.Nodes;
using Aspose.Slides;
using Aspose.Slides.Export;

namespace AsposeMcpServer.Tools;

public class PptMoveSlideTool : IAsposeTool
{
    public string Description => "Move/Reorder a slide within a PowerPoint presentation";

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
            fromIndex = new
            {
                type = "number",
                description = "Current slide index (0-based)"
            },
            toIndex = new
            {
                type = "number",
                description = "Target slide index (0-based)"
            }
        },
        required = new[] { "path", "fromIndex", "toIndex" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var fromIndex = arguments?["fromIndex"]?.GetValue<int>() ?? throw new ArgumentException("fromIndex is required");
        var toIndex = arguments?["toIndex"]?.GetValue<int>() ?? throw new ArgumentException("toIndex is required");

        using var presentation = new Presentation(path);
        var count = presentation.Slides.Count;

        if (fromIndex < 0 || fromIndex >= count)
        {
            throw new ArgumentException($"fromIndex must be between 0 and {count - 1}");
        }
        if (toIndex < 0 || toIndex >= count)
        {
            throw new ArgumentException($"toIndex must be between 0 and {count - 1}");
        }

        var source = presentation.Slides[fromIndex];
        presentation.Slides.InsertClone(toIndex, source);
        var removeIndex = fromIndex + (fromIndex < toIndex ? 1 : 0);
        presentation.Slides.RemoveAt(removeIndex);
        presentation.Save(path, SaveFormat.Pptx);

        return await Task.FromResult($"投影片已移動: {fromIndex} -> {toIndex} (總數 {count})");
    }
}

