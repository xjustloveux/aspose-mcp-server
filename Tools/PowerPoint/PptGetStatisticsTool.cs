using System.Text.Json.Nodes;
using System.Text;
using Aspose.Slides;
using Aspose.Slides.Charts;
using Aspose.Slides.SmartArt;

namespace AsposeMcpServer.Tools;

public class PptGetStatisticsTool : IAsposeTool
{
    public string Description => "Get statistics information about a PowerPoint presentation";

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

        sb.AppendLine("Presentation Statistics:");
        sb.AppendLine($"  Total Slides: {presentation.Slides.Count}");
        sb.AppendLine($"  Total Layouts: {presentation.LayoutSlides.Count}");
        sb.AppendLine($"  Total Masters: {presentation.Masters.Count}");
        sb.AppendLine($"  Slide Size: {presentation.SlideSize.Size.Width} x {presentation.SlideSize.Size.Height}");

        int totalShapes = 0;
        int totalText = 0;
        int totalImages = 0;
        int totalTables = 0;
        int totalCharts = 0;
        int totalSmartArt = 0;
        int totalAudio = 0;
        int totalVideo = 0;
        int totalAnimations = 0;
        int totalHyperlinks = 0;

        foreach (var slide in presentation.Slides)
        {
            totalShapes += slide.Shapes.Count;
            totalAnimations += slide.Timeline.MainSequence.Count;

            foreach (var shape in slide.Shapes)
            {
                if (shape is IAutoShape autoShape && autoShape.TextFrame != null)
                {
                    totalText++;
                    if (!string.IsNullOrWhiteSpace(autoShape.TextFrame.Text))
                    {
                        totalText += autoShape.TextFrame.Text.Length;
                    }
                }
                else if (shape is PictureFrame)
                {
                    totalImages++;
                }
                else if (shape is ITable)
                {
                    totalTables++;
                }
                else if (shape is IChart)
                {
                    totalCharts++;
                }
                else if (shape is ISmartArt)
                {
                    totalSmartArt++;
                }
                else if (shape is IAudioFrame)
                {
                    totalAudio++;
                }
                else if (shape is IVideoFrame)
                {
                    totalVideo++;
                }

                if (shape.HyperlinkClick != null)
                {
                    totalHyperlinks++;
                }
            }
        }

        sb.AppendLine($"  Total Shapes: {totalShapes}");
        sb.AppendLine($"  Total Text Characters: {totalText}");
        sb.AppendLine($"  Total Images: {totalImages}");
        sb.AppendLine($"  Total Tables: {totalTables}");
        sb.AppendLine($"  Total Charts: {totalCharts}");
        sb.AppendLine($"  Total SmartArt: {totalSmartArt}");
        sb.AppendLine($"  Total Audio: {totalAudio}");
        sb.AppendLine($"  Total Video: {totalVideo}");
        sb.AppendLine($"  Total Animations: {totalAnimations}");
        sb.AppendLine($"  Total Hyperlinks: {totalHyperlinks}");

        return await Task.FromResult(sb.ToString());
    }
}

