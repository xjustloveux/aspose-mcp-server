using System.Text.Json.Nodes;
using Aspose.Slides;
using Aspose.Slides.Export;

namespace AsposeMcpServer.Tools;

public class PptAddVideoTool : IAsposeTool
{
    public string Description => "Insert an embedded video frame into a slide";

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
            videoPath = new
            {
                type = "string",
                description = "Video file path"
            },
            x = new
            {
                type = "number",
                description = "X position (optional, default: 50)"
            },
            y = new
            {
                type = "number",
                description = "Y position (optional, default: 50)"
            },
            width = new
            {
                type = "number",
                description = "Width (optional, default: 320)"
            },
            height = new
            {
                type = "number",
                description = "Height (optional, default: 240)"
            }
        },
        required = new[] { "path", "slideIndex", "videoPath" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var slideIndex = arguments?["slideIndex"]?.GetValue<int>() ?? throw new ArgumentException("slideIndex is required");
        var videoPath = arguments?["videoPath"]?.GetValue<string>() ?? throw new ArgumentException("videoPath is required");
        var x = arguments?["x"]?.GetValue<float?>() ?? 50;
        var y = arguments?["y"]?.GetValue<float?>() ?? 50;
        var width = arguments?["width"]?.GetValue<float?>() ?? 320;
        var height = arguments?["height"]?.GetValue<float?>() ?? 240;

        using var presentation = new Presentation(path);
        if (slideIndex < 0 || slideIndex >= presentation.Slides.Count)
        {
            throw new ArgumentException($"slideIndex must be between 0 and {presentation.Slides.Count - 1}");
        }

        var slide = presentation.Slides[slideIndex];
        var frame = slide.Shapes.AddVideoFrame(x, y, width, height, videoPath);

        presentation.Save(path, SaveFormat.Pptx);
        return await Task.FromResult($"已在投影片 {slideIndex} 插入影片: {videoPath}");
    }
}

