using System.Text.Json.Nodes;
using Aspose.Slides;
using Aspose.Slides.Export;

namespace AsposeMcpServer.Tools;

public class PptCopyShapeTool : IAsposeTool
{
    public string Description => "Copy a shape from one slide to another slide";

    public object InputSchema => new
    {
        type = "object",
        properties = new
        {
            path = new { type = "string", description = "Presentation file path" },
            fromSlide = new { type = "number", description = "Source slide index (0-based)" },
            shapeIndex = new { type = "number", description = "Shape index in source slide (0-based)" },
            toSlide = new { type = "number", description = "Target slide index (0-based)" }
        },
        required = new[] { "path", "fromSlide", "shapeIndex", "toSlide" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var fromSlide = arguments?["fromSlide"]?.GetValue<int>() ?? throw new ArgumentException("fromSlide is required");
        var shapeIndex = arguments?["shapeIndex"]?.GetValue<int>() ?? throw new ArgumentException("shapeIndex is required");
        var toSlide = arguments?["toSlide"]?.GetValue<int>() ?? throw new ArgumentException("toSlide is required");

        using var presentation = new Presentation(path);
        if (fromSlide < 0 || fromSlide >= presentation.Slides.Count) throw new ArgumentException("fromSlide out of range");
        if (toSlide < 0 || toSlide >= presentation.Slides.Count) throw new ArgumentException("toSlide out of range");

        var sourceSlide = presentation.Slides[fromSlide];
        if (shapeIndex < 0 || shapeIndex >= sourceSlide.Shapes.Count) throw new ArgumentException("shapeIndex out of range");

        var targetSlide = presentation.Slides[toSlide];
        targetSlide.Shapes.AddClone(sourceSlide.Shapes[shapeIndex]);

        presentation.Save(path, SaveFormat.Pptx);
        return await Task.FromResult($"已複製形狀 {shapeIndex} 從投影片 {fromSlide} 到 {toSlide}");
    }
}

