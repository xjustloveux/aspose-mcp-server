using System.Text.Json.Nodes;
using System.Text;
using Aspose.Slides;
using Aspose.Slides.Export;
using Aspose.Slides.Charts;

namespace AsposeMcpServer.Tools;

public class PptGetShapeDetailsTool : IAsposeTool
{
    public string Description => "Get detailed information about a shape";

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
                description = "Shape index (0-based)"
            }
        },
        required = new[] { "path", "slideIndex", "shapeIndex" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var slideIndex = arguments?["slideIndex"]?.GetValue<int>() ?? throw new ArgumentException("slideIndex is required");
        var shapeIndex = arguments?["shapeIndex"]?.GetValue<int>() ?? throw new ArgumentException("shapeIndex is required");

        using var presentation = new Presentation(path);
        if (slideIndex < 0 || slideIndex >= presentation.Slides.Count)
        {
            throw new ArgumentException($"slideIndex must be between 0 and {presentation.Slides.Count - 1}");
        }

        var slide = presentation.Slides[slideIndex];
        if (shapeIndex < 0 || shapeIndex >= slide.Shapes.Count)
        {
            throw new ArgumentException($"shapeIndex must be between 0 and {slide.Shapes.Count - 1}");
        }

        var shape = slide.Shapes[shapeIndex];
        var sb = new StringBuilder();

        sb.AppendLine($"=== Shape {shapeIndex} Details ===");
        sb.AppendLine($"Type: {shape.GetType().Name}");
        sb.AppendLine($"Position: ({shape.X}, {shape.Y})");
        sb.AppendLine($"Size: ({shape.Width}, {shape.Height})");
        sb.AppendLine($"Rotation: {shape.Rotation}°");
        // Note: Flip properties may not be directly available on IShape

        if (shape is IAutoShape autoShape)
        {
            sb.AppendLine($"\nAutoShape Properties:");
            sb.AppendLine($"  ShapeType: {autoShape.ShapeType}");
            sb.AppendLine($"  Text: {autoShape.TextFrame?.Text ?? "(none)"}");
            if (autoShape.HyperlinkClick != null)
            {
                var url = autoShape.HyperlinkClick.ExternalUrl ?? (autoShape.HyperlinkClick.TargetSlide != null ? $"Slide {presentation.Slides.IndexOf(autoShape.HyperlinkClick.TargetSlide)}" : "Internal link");
                sb.AppendLine($"  Hyperlink: {url}");
            }
        }
        else if (shape is PictureFrame picture)
        {
            sb.AppendLine($"\nPicture Properties:");
            sb.AppendLine($"  AlternativeText: {picture.AlternativeText ?? "(none)"}");
        }
        else if (shape is ITable table)
        {
            sb.AppendLine($"\nTable Properties:");
            sb.AppendLine($"  Rows: {table.Rows.Count}");
            sb.AppendLine($"  Columns: {table.Columns.Count}");
        }
        else if (shape is IChart chart)
        {
            sb.AppendLine($"\nChart Properties:");
            sb.AppendLine($"  ChartType: {chart.Type}");
            sb.AppendLine($"  Title: {chart.ChartTitle?.TextFrameForOverriding?.Text ?? "(none)"}");
        }

        return await Task.FromResult(sb.ToString());
    }
}

