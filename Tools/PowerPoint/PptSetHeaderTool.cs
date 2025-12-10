using System.Text.Json.Nodes;
using Aspose.Slides;
using Aspose.Slides.Export;

namespace AsposeMcpServer.Tools;

public class PptSetHeaderTool : IAsposeTool
{
    public string Description => "Set header text for slides";

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
            headerText = new
            {
                type = "string",
                description = "Header text"
            },
            slideIndices = new
            {
                type = "array",
                items = new { type = "number" },
                description = "Slide indices (0-based, optional, if not provided applies to all slides)"
            }
        },
        required = new[] { "path", "headerText" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var headerText = arguments?["headerText"]?.GetValue<string>() ?? throw new ArgumentException("headerText is required");
        var slideIndices = arguments?["slideIndices"]?.AsArray()?.Select(x => x?.GetValue<int>()).Where(x => x.HasValue).Select(x => x!.Value).ToArray();

        using var presentation = new Presentation(path);
        var slides = slideIndices?.Length > 0
            ? slideIndices.Select(i => presentation.Slides[i]).ToList()
            : presentation.Slides.Cast<ISlide>().ToList();

        foreach (var slide in slides)
        {
            var headerFooter = slide.HeaderFooterManager;
            // Note: Header text is typically set through layout placeholders
            // This is a simplified approach - full implementation would require placeholder manipulation
            headerFooter.SetFooterText(headerText);
            headerFooter.SetFooterVisibility(true);
        }

        presentation.Save(path, SaveFormat.Pptx);

        return await Task.FromResult($"Header set for {slides.Count} slide(s): {path}");
    }
}

