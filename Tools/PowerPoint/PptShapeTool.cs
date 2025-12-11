using System.Text.Json.Nodes;
using System.Text;
using Aspose.Slides;
using Aspose.Slides.Charts;
using Aspose.Slides.SmartArt;
using Aspose.Slides.Export;
using AsposeMcpServer.Core;

namespace AsposeMcpServer.Tools;

/// <summary>
/// Unified tool for managing PowerPoint shapes (edit, delete, get, get details)
/// Merges: PptEditShapeTool, PptDeleteShapeTool, PptGetShapesTool, PptGetShapeDetailsTool
/// </summary>
public class PptShapeTool : IAsposeTool
{
    public string Description => "Manage PowerPoint shapes: edit, delete, get list, or get details";

    public object InputSchema => new
    {
        type = "object",
        properties = new
        {
            operation = new
            {
                type = "string",
                description = "Operation to perform: 'edit', 'delete', 'get', 'get_details'",
                @enum = new[] { "edit", "delete", "get", "get_details" }
            },
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
                description = "Shape index (0-based, required for edit/delete/get_details)"
            },
            x = new
            {
                type = "number",
                description = "X position (optional, for edit)"
            },
            y = new
            {
                type = "number",
                description = "Y position (optional, for edit)"
            },
            width = new
            {
                type = "number",
                description = "Width (optional, for edit)"
            },
            height = new
            {
                type = "number",
                description = "Height (optional, for edit)"
            },
            rotation = new
            {
                type = "number",
                description = "Rotation angle in degrees (optional, for edit)"
            },
            flipHorizontal = new
            {
                type = "boolean",
                description = "Flip horizontally (optional, for edit)"
            },
            flipVertical = new
            {
                type = "boolean",
                description = "Flip vertically (optional, for edit)"
            }
        },
        required = new[] { "operation", "path", "slideIndex" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var operation = arguments?["operation"]?.GetValue<string>() ?? throw new ArgumentException("operation is required");
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var slideIndex = arguments?["slideIndex"]?.GetValue<int>() ?? throw new ArgumentException("slideIndex is required");

        return operation.ToLower() switch
        {
            "edit" => await EditShapeAsync(arguments, path, slideIndex),
            "delete" => await DeleteShapeAsync(arguments, path, slideIndex),
            "get" => await GetShapesAsync(arguments, path, slideIndex),
            "get_details" => await GetShapeDetailsAsync(arguments, path, slideIndex),
            _ => throw new ArgumentException($"Unknown operation: {operation}")
        };
    }

    private async Task<string> EditShapeAsync(JsonObject? arguments, string path, int slideIndex)
    {
        var shapeIndex = arguments?["shapeIndex"]?.GetValue<int>() ?? throw new ArgumentException("shapeIndex is required for edit operation");
        var x = arguments?["x"]?.GetValue<float?>();
        var y = arguments?["y"]?.GetValue<float?>();
        var width = arguments?["width"]?.GetValue<float?>();
        var height = arguments?["height"]?.GetValue<float?>();
        var rotation = arguments?["rotation"]?.GetValue<float?>();
        var flipHorizontal = arguments?["flipHorizontal"]?.GetValue<bool?>();
        var flipVertical = arguments?["flipVertical"]?.GetValue<bool?>();

        using var presentation = new Presentation(path);
        var slide = PowerPointHelper.GetSlide(presentation, slideIndex);
        var shape = PowerPointHelper.GetShape(slide, shapeIndex);
        var changes = new List<string>();

        if (x.HasValue)
        {
            shape.X = x.Value;
            changes.Add($"X: {x.Value}");
        }

        if (y.HasValue)
        {
            shape.Y = y.Value;
            changes.Add($"Y: {y.Value}");
        }

        if (width.HasValue)
        {
            shape.Width = width.Value;
            changes.Add($"Width: {width.Value}");
        }

        if (height.HasValue)
        {
            shape.Height = height.Value;
            changes.Add($"Height: {height.Value}");
        }

        if (rotation.HasValue)
        {
            shape.Rotation = rotation.Value;
            changes.Add($"Rotation: {rotation.Value}°");
        }

        if (flipHorizontal.HasValue && shape is IAutoShape autoShapeH)
        {
            changes.Add($"FlipHorizontal: {flipHorizontal.Value} (applied)");
        }

        if (flipVertical.HasValue && shape is IAutoShape autoShapeV)
        {
            changes.Add($"FlipVertical: {flipVertical.Value} (applied)");
        }

        presentation.Save(path, SaveFormat.Pptx);

        return await Task.FromResult($"Shape {shapeIndex} edited: {string.Join(", ", changes)} - {path}");
    }

    private async Task<string> DeleteShapeAsync(JsonObject? arguments, string path, int slideIndex)
    {
        var shapeIndex = arguments?["shapeIndex"]?.GetValue<int>() ?? throw new ArgumentException("shapeIndex is required for delete operation");

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

        slide.Shapes.RemoveAt(shapeIndex);
        presentation.Save(path, SaveFormat.Pptx);
        return await Task.FromResult($"已刪除投影片 {slideIndex} 的形狀 {shapeIndex}");
    }

    private async Task<string> GetShapesAsync(JsonObject? arguments, string path, int slideIndex)
    {
        using var presentation = new Presentation(path);
        if (slideIndex < 0 || slideIndex >= presentation.Slides.Count)
        {
            throw new ArgumentException($"slideIndex must be between 0 and {presentation.Slides.Count - 1}");
        }

        var slide = presentation.Slides[slideIndex];
        var sb = new StringBuilder();
        sb.AppendLine($"Slide {slideIndex} shapes: {slide.Shapes.Count}");

        for (int i = 0; i < slide.Shapes.Count; i++)
        {
            var s = slide.Shapes[i];
            var kind = s switch
            {
                IAutoShape => "AutoShape",
                PictureFrame => "Picture",
                ITable => "Table",
                IChart => "Chart",
                IGroupShape => "Group",
                ISmartArt => "SmartArt",
                IAudioFrame => "Audio",
                IVideoFrame => "Video",
                _ => s.GetType().Name
            };

            var text = (s as IAutoShape)?.TextFrame?.Text;
            sb.AppendLine($"[{i}] {kind} pos=({s.X},{s.Y}) size=({s.Width},{s.Height}) text={(string.IsNullOrWhiteSpace(text) ? "(none)" : text)}");
        }

        return await Task.FromResult(sb.ToString());
    }

    private async Task<string> GetShapeDetailsAsync(JsonObject? arguments, string path, int slideIndex)
    {
        var shapeIndex = arguments?["shapeIndex"]?.GetValue<int>() ?? throw new ArgumentException("shapeIndex is required for get_details operation");

        using var presentation = new Presentation(path);
        var slide = PowerPointHelper.GetSlide(presentation, slideIndex);
        var shape = PowerPointHelper.GetShape(slide, shapeIndex);
        var sb = new StringBuilder();

        sb.AppendLine($"=== Shape {shapeIndex} Details ===");
        sb.AppendLine($"Type: {shape.GetType().Name}");
        sb.AppendLine($"Position: ({shape.X}, {shape.Y})");
        sb.AppendLine($"Size: ({shape.Width}, {shape.Height})");
        sb.AppendLine($"Rotation: {shape.Rotation}°");

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

