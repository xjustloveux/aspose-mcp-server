using System.Text.Json.Nodes;
using Aspose.Slides;
using Aspose.Slides.Export;

namespace AsposeMcpServer.Tools;

public class PptReorderShapeTool : IAsposeTool
{
    public string Description => "Reorder a shape (bring forward/backward) by moving to target index";

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
            shapeIndex = new
            {
                type = "number",
                description = "Shape index to move (0-based)"
            },
            toIndex = new
            {
                type = "number",
                description = "Target index (0-based)"
            }
        },
        required = new[] { "path", "slideIndex", "shapeIndex", "toIndex" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var slideIndex = arguments?["slideIndex"]?.GetValue<int>() ?? throw new ArgumentException("slideIndex is required");
        var shapeIndex = arguments?["shapeIndex"]?.GetValue<int>() ?? throw new ArgumentException("shapeIndex is required");
        var toIndex = arguments?["toIndex"]?.GetValue<int>() ?? throw new ArgumentException("toIndex is required");

        using var presentation = new Presentation(path);
        if (slideIndex < 0 || slideIndex >= presentation.Slides.Count)
        {
            throw new ArgumentException($"slideIndex must be between 0 and {presentation.Slides.Count - 1}");
        }

        var slide = presentation.Slides[slideIndex];
        var count = slide.Shapes.Count;
        if (shapeIndex < 0 || shapeIndex >= count) throw new ArgumentException($"shapeIndex must be between 0 and {count - 1}");
        if (toIndex < 0 || toIndex >= count) throw new ArgumentException($"toIndex must be between 0 and {count - 1}");

        var shape = slide.Shapes[shapeIndex];
        slide.Shapes.InsertClone(toIndex, shape);
        var removeIndex = shapeIndex + (shapeIndex < toIndex ? 1 : 0);
        slide.Shapes.RemoveAt(removeIndex);

        presentation.Save(path, SaveFormat.Pptx);
        return await Task.FromResult($"已移動形狀 {shapeIndex} -> {toIndex}");
    }
}

