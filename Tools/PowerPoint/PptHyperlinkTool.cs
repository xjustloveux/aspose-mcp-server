using System.Drawing;
using System.Text;
using System.Text.Json.Nodes;
using Aspose.Slides;
using Aspose.Slides.Export;
using AsposeMcpServer.Core;

namespace AsposeMcpServer.Tools.PowerPoint;

/// <summary>
///     Unified tool for managing PowerPoint hyperlinks (add, edit, delete, get)
///     Merges: PptAddHyperlinkTool, PptEditHyperlinkTool, PptDeleteHyperlinkTool, PptGetHyperlinksTool
/// </summary>
public class PptHyperlinkTool : IAsposeTool
{
    public string Description => @"Manage PowerPoint hyperlinks. Supports 4 operations: add, edit, delete, get.

Usage examples:
- Add hyperlink: ppt_hyperlink(operation='add', path='presentation.pptx', slideIndex=0, text='Link', url='https://example.com')
- Edit hyperlink: ppt_hyperlink(operation='edit', path='presentation.pptx', slideIndex=0, shapeIndex=0, url='https://newurl.com')
- Delete hyperlink: ppt_hyperlink(operation='delete', path='presentation.pptx', slideIndex=0, shapeIndex=0)
- Get hyperlinks: ppt_hyperlink(operation='get', path='presentation.pptx', slideIndex=0)";

    public object InputSchema => new
    {
        type = "object",
        properties = new
        {
            operation = new
            {
                type = "string",
                description = @"Operation to perform.
- 'add': Add hyperlink to shape (required params: path, slideIndex, text, url)
- 'edit': Edit hyperlink (required params: path, slideIndex, shapeIndex, url)
- 'delete': Delete hyperlink (required params: path, slideIndex, shapeIndex)
- 'get': Get all hyperlinks (required params: path, slideIndex)",
                @enum = new[] { "add", "edit", "delete", "get" }
            },
            path = new
            {
                type = "string",
                description = "Presentation file path (required for all operations)"
            },
            slideIndex = new
            {
                type = "number",
                description = "Slide index (0-based)"
            },
            shapeIndex = new
            {
                type = "number",
                description = "Shape index (0-based, required for edit/delete, optional for add)"
            },
            text = new
            {
                type = "string",
                description = "Display text (required for add)"
            },
            url = new
            {
                type = "string",
                description = "Hyperlink URL (required for add, optional for edit)"
            },
            slideTargetIndex = new
            {
                type = "number",
                description = "Target slide index for internal link (optional, for edit)"
            },
            removeHyperlink = new
            {
                type = "boolean",
                description = "Remove hyperlink (optional, for edit)"
            },
            x = new
            {
                type = "number",
                description = "X position (optional, for add, default: 50)"
            },
            y = new
            {
                type = "number",
                description = "Y position (optional, for add, default: 50)"
            },
            width = new
            {
                type = "number",
                description = "Width (optional, for add, default: 300)"
            },
            height = new
            {
                type = "number",
                description = "Height (optional, for add, default: 50)"
            },
            outputPath = new
            {
                type = "string",
                description = "Output file path (optional, for add/edit/delete operations, defaults to input path)"
            }
        },
        required = new[] { "operation", "path" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var operation = ArgumentHelper.GetString(arguments, "operation");
        var path = ArgumentHelper.GetAndValidatePath(arguments);

        return operation.ToLower() switch
        {
            "add" => await AddHyperlinkAsync(arguments, path),
            "edit" => await EditHyperlinkAsync(arguments, path),
            "delete" => await DeleteHyperlinkAsync(arguments, path),
            "get" => await GetHyperlinksAsync(arguments, path),
            _ => throw new ArgumentException($"Unknown operation: {operation}")
        };
    }

    /// <summary>
    ///     Adds a hyperlink to a shape
    /// </summary>
    /// <param name="arguments">JSON arguments containing slideIndex, shapeIndex, address, optional text, outputPath</param>
    /// <param name="path">PowerPoint file path</param>
    /// <returns>Success message</returns>
    private async Task<string> AddHyperlinkAsync(JsonObject? arguments, string path)
    {
        var slideIndex = ArgumentHelper.GetInt(arguments, "slideIndex");
        var text = ArgumentHelper.GetString(arguments, "text");
        var url = ArgumentHelper.GetString(arguments, "url");
        var x = ArgumentHelper.GetFloat(arguments, "x", 50);
        var y = ArgumentHelper.GetFloat(arguments, "y", 50);
        var width = ArgumentHelper.GetFloat(arguments, "width", 300);
        var height = ArgumentHelper.GetFloat(arguments, "height", 50);

        using var presentation = new Presentation(path);
        if (slideIndex < 0 || slideIndex >= presentation.Slides.Count)
            throw new ArgumentException($"slideIndex must be between 0 and {presentation.Slides.Count - 1}");

        var slide = presentation.Slides[slideIndex];
        var shape = slide.Shapes.AddAutoShape(ShapeType.Rectangle, x, y, width, height);
        var textFrame = shape.TextFrame;
        textFrame.Text = text;

        var portion = textFrame.Paragraphs[0].Portions[0];
        portion.PortionFormat.HyperlinkClick = new Hyperlink(url);
        portion.PortionFormat.FontHeight = 14;
        portion.PortionFormat.FillFormat.FillType = FillType.Solid;
        portion.PortionFormat.FillFormat.SolidFillColor.Color = Color.Blue;

        var outputPath = ArgumentHelper.GetAndValidateOutputPath(arguments, path);
        presentation.Save(outputPath, SaveFormat.Pptx);
        return await Task.FromResult($"Hyperlink added to slide {slideIndex}: {url}");
    }

    /// <summary>
    ///     Edits an existing hyperlink
    /// </summary>
    /// <param name="arguments">JSON arguments containing slideIndex, shapeIndex, optional address, text, outputPath</param>
    /// <param name="path">PowerPoint file path</param>
    /// <returns>Success message</returns>
    private async Task<string> EditHyperlinkAsync(JsonObject? arguments, string path)
    {
        var slideIndex = ArgumentHelper.GetInt(arguments, "slideIndex");
        var shapeIndex = ArgumentHelper.GetInt(arguments, "shapeIndex");
        var url = ArgumentHelper.GetStringNullable(arguments, "url");
        var slideTargetIndex = ArgumentHelper.GetIntNullable(arguments, "slideTargetIndex");
        var removeHyperlink = ArgumentHelper.GetBool(arguments, "removeHyperlink", false);

        using var presentation = new Presentation(path);
        var slide = PowerPointHelper.GetSlide(presentation, slideIndex);
        var shape = PowerPointHelper.GetShape(slide, shapeIndex);

        if (removeHyperlink)
        {
            if (shape is IAutoShape { TextFrame: not null } autoShape)
                foreach (var paragraph in autoShape.TextFrame.Paragraphs)
                foreach (var portion in paragraph.Portions)
                    portion.PortionFormat.HyperlinkClick = null;

            shape.HyperlinkClick = null;
        }
        else if (!string.IsNullOrEmpty(url))
        {
            shape.HyperlinkClick = new Hyperlink(url);
        }
        else if (slideTargetIndex.HasValue)
        {
            if (slideTargetIndex.Value < 0 || slideTargetIndex.Value >= presentation.Slides.Count)
                throw new ArgumentException($"slideTargetIndex must be between 0 and {presentation.Slides.Count - 1}");
            shape.HyperlinkClick = new Hyperlink(presentation.Slides[slideTargetIndex.Value]);
        }
        else
        {
            throw new ArgumentException("Either url, slideTargetIndex, or removeHyperlink must be provided");
        }

        var outputPath = ArgumentHelper.GetAndValidateOutputPath(arguments, path);
        presentation.Save(outputPath, SaveFormat.Pptx);
        return await Task.FromResult($"Hyperlink updated on slide {slideIndex}, shape {shapeIndex}");
    }

    /// <summary>
    ///     Deletes a hyperlink from a shape
    /// </summary>
    /// <param name="arguments">JSON arguments containing slideIndex, shapeIndex, optional outputPath</param>
    /// <param name="path">PowerPoint file path</param>
    /// <returns>Success message</returns>
    private async Task<string> DeleteHyperlinkAsync(JsonObject? arguments, string path)
    {
        var slideIndex = ArgumentHelper.GetInt(arguments, "slideIndex");
        var shapeIndex = ArgumentHelper.GetInt(arguments, "shapeIndex");

        using var presentation = new Presentation(path);
        var slide = PowerPointHelper.GetSlide(presentation, slideIndex);
        var shape = PowerPointHelper.GetShape(slide, shapeIndex);
        shape.HyperlinkClick = null;

        if (shape is IAutoShape { TextFrame: not null } autoShape)
            foreach (var paragraph in autoShape.TextFrame.Paragraphs)
            foreach (var portion in paragraph.Portions)
                portion.PortionFormat.HyperlinkClick = null;

        var outputPath = ArgumentHelper.GetAndValidateOutputPath(arguments, path);
        presentation.Save(outputPath, SaveFormat.Pptx);
        return await Task.FromResult($"Hyperlink deleted from slide {slideIndex}, shape {shapeIndex}");
    }

    /// <summary>
    ///     Gets all hyperlinks from the presentation
    /// </summary>
    /// <param name="arguments">JSON arguments (no specific parameters required)</param>
    /// <param name="path">PowerPoint file path</param>
    /// <returns>Formatted string with all hyperlinks</returns>
    private async Task<string> GetHyperlinksAsync(JsonObject? arguments, string path)
    {
        var slideIndex = ArgumentHelper.GetIntNullable(arguments, "slideIndex");

        using var presentation = new Presentation(path);
        var sb = new StringBuilder();

        if (slideIndex.HasValue)
        {
            if (slideIndex.Value < 0 || slideIndex.Value >= presentation.Slides.Count)
                throw new ArgumentException($"slideIndex must be between 0 and {presentation.Slides.Count - 1}");
            var slide = presentation.Slides[slideIndex.Value];
            sb.AppendLine($"=== Slide {slideIndex.Value} Hyperlinks ===");
            GetHyperlinksFromSlide(presentation, slide, sb);
        }
        else
        {
            sb.AppendLine("=== All Hyperlinks ===");
            for (var i = 0; i < presentation.Slides.Count; i++)
            {
                var slide = presentation.Slides[i];
                var hyperlinks = GetHyperlinksFromSlide(presentation, slide, null);
                if (hyperlinks > 0)
                {
                    sb.AppendLine($"\nSlide {i}: {hyperlinks} hyperlink(s)");
                    GetHyperlinksFromSlide(presentation, slide, sb);
                }
            }
        }

        return await Task.FromResult(sb.ToString());
    }

    private int GetHyperlinksFromSlide(IPresentation presentation, ISlide slide, StringBuilder? sb)
    {
        var count = 0;
        foreach (var shape in slide.Shapes)
            if (shape is IAutoShape autoShape)
            {
                if (autoShape.HyperlinkClick != null)
                {
                    count++;
                    if (sb != null)
                    {
                        var url = autoShape.HyperlinkClick.ExternalUrl ?? (autoShape.HyperlinkClick.TargetSlide != null
                            ? $"Slide {presentation.Slides.IndexOf(autoShape.HyperlinkClick.TargetSlide)}"
                            : "Internal link");
                        sb.AppendLine($"  Shape [{slide.Shapes.IndexOf(shape)}]: {url}");
                    }
                }

                if (autoShape.HyperlinkMouseOver != null)
                {
                    count++;
                    if (sb != null)
                    {
                        var url = autoShape.HyperlinkMouseOver.ExternalUrl ??
                                  (autoShape.HyperlinkMouseOver.TargetSlide != null
                                      ? $"Slide {presentation.Slides.IndexOf(autoShape.HyperlinkMouseOver.TargetSlide)}"
                                      : "Internal link");
                        sb.AppendLine($"  Shape [{slide.Shapes.IndexOf(shape)}] (mouseover): {url}");
                    }
                }
            }

        return count;
    }
}