using System.Text.Json;
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

    /// <summary>
    ///     Executes the tool operation with the provided JSON arguments
    /// </summary>
    /// <param name="arguments">JSON arguments object containing operation parameters</param>
    /// <returns>Result message as a string</returns>
    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var operation = ArgumentHelper.GetString(arguments, "operation");
        var path = ArgumentHelper.GetAndValidatePath(arguments);
        var outputPath = ArgumentHelper.GetAndValidateOutputPath(arguments, path);

        return operation.ToLower() switch
        {
            "add" => await AddHyperlinkAsync(path, outputPath, arguments),
            "edit" => await EditHyperlinkAsync(path, outputPath, arguments),
            "delete" => await DeleteHyperlinkAsync(path, outputPath, arguments),
            "get" => await GetHyperlinksAsync(path, arguments),
            _ => throw new ArgumentException($"Unknown operation: {operation}")
        };
    }

    /// <summary>
    ///     Adds a hyperlink to a shape
    /// </summary>
    /// <param name="path">PowerPoint file path</param>
    /// <param name="outputPath">Output file path</param>
    /// <param name="arguments">JSON arguments containing slideIndex, shapeIndex, address, optional text</param>
    /// <returns>Success message</returns>
    private Task<string> AddHyperlinkAsync(string path, string outputPath, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var slideIndex = ArgumentHelper.GetInt(arguments, "slideIndex");
            var url = ArgumentHelper.GetString(arguments, "url");
            var displayText = ArgumentHelper.GetStringNullable(arguments, "displayText");
            var shapeIndex = ArgumentHelper.GetIntNullable(arguments, "shapeIndex");
            var x = ArgumentHelper.GetFloatNullable(arguments, "x");
            var y = ArgumentHelper.GetFloatNullable(arguments, "y");
            var width = ArgumentHelper.GetFloatNullable(arguments, "width");
            var height = ArgumentHelper.GetFloatNullable(arguments, "height");

            using var presentation = new Presentation(path);
            if (slideIndex < 0 || slideIndex >= presentation.Slides.Count)
                throw new ArgumentException($"slideIndex must be between 0 and {presentation.Slides.Count - 1}");

            var slide = presentation.Slides[slideIndex];
            IShape shape;

            if (shapeIndex is >= 0 && shapeIndex.Value < slide.Shapes.Count)
            {
                // Use existing shape
                shape = slide.Shapes[shapeIndex.Value];
            }
            else
            {
                // Create new shape
                var defaultX = x ?? 50;
                var defaultY = y ?? 50;
                var defaultWidth = width ?? 300;
                var defaultHeight = height ?? 50;
                shape = slide.Shapes.AddAutoShape(ShapeType.Rectangle, defaultX, defaultY, defaultWidth, defaultHeight);
            }

            // Set hyperlink on shape
            shape.HyperlinkClick = new Hyperlink(url);

            // Also set display text if provided
            if (!string.IsNullOrEmpty(displayText) && shape is IAutoShape { TextFrame: not null } autoShape)
                autoShape.TextFrame.Text = displayText;

            presentation.Save(outputPath, SaveFormat.Pptx);
            return $"Hyperlink added to slide {slideIndex}: {url}. Output: {outputPath}";
        });
    }

    /// <summary>
    ///     Edits an existing hyperlink
    /// </summary>
    /// <param name="path">PowerPoint file path</param>
    /// <param name="outputPath">Output file path</param>
    /// <param name="arguments">JSON arguments containing slideIndex, shapeIndex, optional address, text</param>
    /// <returns>Success message</returns>
    private Task<string> EditHyperlinkAsync(string path, string outputPath, JsonObject? arguments)
    {
        return Task.Run(() =>
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
                    throw new ArgumentException(
                        $"slideTargetIndex must be between 0 and {presentation.Slides.Count - 1}");
                shape.HyperlinkClick = new Hyperlink(presentation.Slides[slideTargetIndex.Value]);
            }
            else
            {
                throw new ArgumentException("Either url, slideTargetIndex, or removeHyperlink must be provided");
            }

            presentation.Save(outputPath, SaveFormat.Pptx);
            return $"Hyperlink updated on slide {slideIndex}, shape {shapeIndex}. Output: {outputPath}";
        });
    }

    /// <summary>
    ///     Deletes a hyperlink from a shape
    /// </summary>
    /// <param name="path">PowerPoint file path</param>
    /// <param name="outputPath">Output file path</param>
    /// <param name="arguments">JSON arguments containing slideIndex, shapeIndex</param>
    /// <returns>Success message</returns>
    private Task<string> DeleteHyperlinkAsync(string path, string outputPath, JsonObject? arguments)
    {
        return Task.Run(() =>
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

            presentation.Save(outputPath, SaveFormat.Pptx);
            return $"Hyperlink deleted from slide {slideIndex}, shape {shapeIndex}. Output: {outputPath}";
        });
    }

    /// <summary>
    ///     Gets all hyperlinks from the presentation
    /// </summary>
    /// <param name="path">PowerPoint file path</param>
    /// <param name="arguments">JSON arguments (no specific parameters required)</param>
    /// <returns>JSON string with all hyperlinks</returns>
    private Task<string> GetHyperlinksAsync(string path, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var slideIndex = ArgumentHelper.GetIntNullable(arguments, "slideIndex");

            using var presentation = new Presentation(path);

            if (slideIndex.HasValue)
            {
                if (slideIndex.Value < 0 || slideIndex.Value >= presentation.Slides.Count)
                    throw new ArgumentException($"slideIndex must be between 0 and {presentation.Slides.Count - 1}");
                var slide = presentation.Slides[slideIndex.Value];
                var hyperlinksList = GetHyperlinksFromSlideAsJson(presentation, slide);

                var result = new
                {
                    slideIndex = slideIndex.Value,
                    count = hyperlinksList.Count,
                    hyperlinks = hyperlinksList
                };

                return JsonSerializer.Serialize(result, new JsonSerializerOptions { WriteIndented = true });
            }
            else
            {
                var slidesList = new List<object>();
                var totalCount = 0;

                for (var i = 0; i < presentation.Slides.Count; i++)
                {
                    var slide = presentation.Slides[i];
                    var hyperlinksList = GetHyperlinksFromSlideAsJson(presentation, slide);
                    totalCount += hyperlinksList.Count;

                    slidesList.Add(new
                    {
                        slideIndex = i,
                        count = hyperlinksList.Count,
                        hyperlinks = hyperlinksList
                    });
                }

                var result = new
                {
                    totalCount,
                    slides = slidesList
                };

                return JsonSerializer.Serialize(result, new JsonSerializerOptions { WriteIndented = true });
            }
        });
    }

    private List<object> GetHyperlinksFromSlideAsJson(IPresentation presentation, ISlide slide)
    {
        var hyperlinksList = new List<object>();

        foreach (var shape in slide.Shapes)
            if (shape is IAutoShape autoShape)
            {
                if (autoShape.HyperlinkClick != null)
                {
                    var url = autoShape.HyperlinkClick.ExternalUrl ?? (autoShape.HyperlinkClick.TargetSlide != null
                        ? $"Slide {presentation.Slides.IndexOf(autoShape.HyperlinkClick.TargetSlide)}"
                        : "Internal link");

                    hyperlinksList.Add(new
                    {
                        shapeIndex = slide.Shapes.IndexOf(shape),
                        triggerType = "click",
                        url
                    });
                }

                if (autoShape.HyperlinkMouseOver != null)
                {
                    var url = autoShape.HyperlinkMouseOver.ExternalUrl ??
                              (autoShape.HyperlinkMouseOver.TargetSlide != null
                                  ? $"Slide {presentation.Slides.IndexOf(autoShape.HyperlinkMouseOver.TargetSlide)}"
                                  : "Internal link");

                    hyperlinksList.Add(new
                    {
                        shapeIndex = slide.Shapes.IndexOf(shape),
                        triggerType = "mouseover",
                        url
                    });
                }
            }

        return hyperlinksList;
    }
}