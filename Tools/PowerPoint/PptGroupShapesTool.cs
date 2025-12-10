using System.Text.Json.Nodes;
using Aspose.Slides;
using Aspose.Slides.Export;

namespace AsposeMcpServer.Tools;

public class PptGroupShapesTool : IAsposeTool
{
    public string Description => "Group multiple shapes together on a PowerPoint slide";

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
            shapeIndices = new
            {
                type = "array",
                items = new { type = "number" },
                description = "Array of shape indices to group (must have at least 2 shapes)"
            }
        },
        required = new[] { "path", "slideIndex", "shapeIndices" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var slideIndex = arguments?["slideIndex"]?.GetValue<int>() ?? throw new ArgumentException("slideIndex is required");
        var shapeIndicesArray = arguments?["shapeIndices"]?.AsArray() ?? throw new ArgumentException("shapeIndices is required");

        var shapeIndices = shapeIndicesArray.Select(s => s?.GetValue<int>()).Where(s => s.HasValue).Select(s => s!.Value).OrderByDescending(s => s).ToList();

        if (shapeIndices.Count < 2)
        {
            throw new ArgumentException("At least 2 shapes are required for grouping");
        }

        using var presentation = new Presentation(path);
        if (slideIndex < 0 || slideIndex >= presentation.Slides.Count)
        {
            throw new ArgumentException($"slideIndex must be between 0 and {presentation.Slides.Count - 1}");
        }

        var slide = presentation.Slides[slideIndex];
        var shapesToGroup = new List<IShape>();

        foreach (var idx in shapeIndices)
        {
            if (idx < 0 || idx >= slide.Shapes.Count)
            {
                throw new ArgumentException($"shapeIndex {idx} is out of range");
            }
            shapesToGroup.Add(slide.Shapes[idx]);
        }

        // Group shapes - create a group shape and add shapes to it
        var groupShape = slide.Shapes.AddGroupShape();
        foreach (var shape in shapesToGroup)
        {
            slide.Shapes.Remove(shape);
            groupShape.Shapes.AddClone(shape);
        }

        presentation.Save(path, SaveFormat.Pptx);
        return await Task.FromResult($"Grouped {shapeIndices.Count} shapes on slide {slideIndex}");
    }
}

