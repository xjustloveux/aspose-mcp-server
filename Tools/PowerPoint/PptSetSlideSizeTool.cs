using System.Text.Json.Nodes;
using Aspose.Slides;
using Aspose.Slides.Export;

namespace AsposeMcpServer.Tools;

public class PptSetSlideSizeTool : IAsposeTool
{
    public string Description => "Set slide size by preset or custom width/height";

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
            preset = new
            {
                type = "string",
                description = "Preset: OnScreen16x9, OnScreen16x10, Letter, A4, Banner, Custom"
            },
            width = new
            {
                type = "number",
                description = "Custom width (points) when preset=Custom"
            },
            height = new
            {
                type = "number",
                description = "Custom height (points) when preset=Custom"
            }
        },
        required = new[] { "path" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var preset = arguments?["preset"]?.GetValue<string>() ?? "OnScreen16x9";
        var width = arguments?["width"]?.GetValue<double?>();
        var height = arguments?["height"]?.GetValue<double?>();

        using var presentation = new Presentation(path);
        var slideSize = presentation.SlideSize;
        var type = preset.ToLower() switch
        {
            "onscreen16x10" => SlideSizeType.OnScreen16x10,
            "a4" => SlideSizeType.A4Paper,
            "banner" => SlideSizeType.Banner,
            "custom" => SlideSizeType.Custom,
            _ => SlideSizeType.OnScreen
        };

        if (type == SlideSizeType.Custom)
        {
            if (!width.HasValue || !height.HasValue)
            {
                throw new ArgumentException("custom size requires width and height");
            }
            slideSize.SetSize((float)width.Value, (float)height.Value, SlideSizeScaleType.DoNotScale);
        }
        else
        {
            slideSize.SetSize(type, SlideSizeScaleType.DoNotScale);
        }

        presentation.Save(path, SaveFormat.Pptx);
        return await Task.FromResult($"已設定投影片尺寸: {slideSize.Type} {(slideSize.Type == SlideSizeType.Custom ? $"{slideSize.Size.Width}x{slideSize.Size.Height}" : string.Empty)}");
    }
}

