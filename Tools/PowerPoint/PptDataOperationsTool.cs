using System.Text.Json.Nodes;
using System.Text;
using Aspose.Slides;
using Aspose.Slides.Charts;
using Aspose.Slides.SmartArt;
using Aspose.Slides.Export;
using AsposeMcpServer.Core;

namespace AsposeMcpServer.Tools;

/// <summary>
/// Unified tool for PowerPoint data operations (get statistics, get content, get slide details)
/// Merges: PptGetStatisticsTool, PptGetContentTool, PptGetSlideDetailsTool
/// </summary>
public class PptDataOperationsTool : IAsposeTool
{
    public string Description => "PowerPoint data operations: get statistics, get content, or get slide details";

    public object InputSchema => new
    {
        type = "object",
        properties = new
        {
            operation = new
            {
                type = "string",
                description = "Operation to perform: 'get_statistics', 'get_content', 'get_slide_details'",
                @enum = new[] { "get_statistics", "get_content", "get_slide_details" }
            },
            path = new
            {
                type = "string",
                description = "Presentation file path"
            },
            slideIndex = new
            {
                type = "number",
                description = "Slide index (0-based, required for get_slide_details)"
            }
        },
        required = new[] { "operation", "path" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var operation = arguments?["operation"]?.GetValue<string>() ?? throw new ArgumentException("operation is required");
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");

        return operation.ToLower() switch
        {
            "get_statistics" => await GetStatisticsAsync(arguments, path),
            "get_content" => await GetContentAsync(arguments, path),
            "get_slide_details" => await GetSlideDetailsAsync(arguments, path),
            _ => throw new ArgumentException($"Unknown operation: {operation}")
        };
    }

    private async Task<string> GetStatisticsAsync(JsonObject? arguments, string path)
    {
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

    private async Task<string> GetContentAsync(JsonObject? arguments, string path)
    {
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

    private async Task<string> GetSlideDetailsAsync(JsonObject? arguments, string path)
    {
        var slideIndex = arguments?["slideIndex"]?.GetValue<int>() ?? throw new ArgumentException("slideIndex is required for get_slide_details operation");

        using var presentation = new Presentation(path);
        var slide = PowerPointHelper.GetSlide(presentation, slideIndex);
        var sb = new StringBuilder();

        sb.AppendLine($"=== Slide {slideIndex} Details ===");
        sb.AppendLine($"Hidden: {slide.Hidden}");
        sb.AppendLine($"Layout: {slide.LayoutSlide?.Name ?? "None"}");
        sb.AppendLine($"Shapes Count: {slide.Shapes.Count}");

        // Transition
        var transition = slide.SlideShowTransition;
        if (transition != null)
        {
            sb.AppendLine($"\nTransition:");
            sb.AppendLine($"  Type: {transition.Type}");
            sb.AppendLine($"  Speed: {transition.Speed}");
            sb.AppendLine($"  AdvanceOnClick: {transition.AdvanceOnClick}");
            sb.AppendLine($"  AdvanceAfterTime: {transition.AdvanceAfterTime}ms");
        }

        // Animations
        var animations = slide.Timeline.MainSequence;
        sb.AppendLine($"\nAnimations: {animations.Count}");
        for (int i = 0; i < animations.Count; i++)
        {
            var anim = animations[i];
            sb.AppendLine($"  [{i}] Type: {anim.Type}, Shape: {anim.TargetShape?.GetType().Name}");
        }

        // Background
        var background = slide.Background;
        if (background != null)
        {
            sb.AppendLine($"\nBackground:");
            sb.AppendLine($"  FillType: {background.FillFormat.FillType}");
        }

        // Notes
        var notesSlide = slide.NotesSlideManager.NotesSlide;
        if (notesSlide != null)
        {
            var notesText = notesSlide.NotesTextFrame?.Text;
            sb.AppendLine($"\nNotes: {(string.IsNullOrWhiteSpace(notesText) ? "None" : notesText)}");
        }

        return await Task.FromResult(sb.ToString());
    }
}

