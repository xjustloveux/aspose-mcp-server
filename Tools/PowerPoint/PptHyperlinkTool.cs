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
- Add hyperlink (URL, shape-level): ppt_hyperlink(operation='add', path='presentation.pptx', slideIndex=0, text='Click here', url='https://example.com')
- Add hyperlink (URL, text-level): ppt_hyperlink(operation='add', path='presentation.pptx', slideIndex=0, text='Please click here for more info', linkText='here', url='https://example.com')
- Add hyperlink (internal): ppt_hyperlink(operation='add', path='presentation.pptx', slideIndex=0, text='Go to slide 5', slideTargetIndex=4)
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
- 'add': Add hyperlink to shape (required params: path, slideIndex, text, url or slideTargetIndex)
- 'edit': Edit hyperlink (required params: path, slideIndex, shapeIndex, url or slideTargetIndex)
- 'delete': Delete hyperlink (required params: path, slideIndex, shapeIndex)
- 'get': Get all hyperlinks (required params: path)",
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
            linkText = new
            {
                type = "string",
                description =
                    "Specific text to apply hyperlink to (optional, for add). When provided, only this text portion will have the hyperlink. When omitted, the entire shape will be clickable."
            },
            url = new
            {
                type = "string",
                description = "Hyperlink URL (required for add, optional for edit)"
            },
            slideTargetIndex = new
            {
                type = "number",
                description = "Target slide index for internal link (0-based, optional, for add/edit)"
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
    ///     Executes the tool operation with the provided JSON arguments.
    /// </summary>
    /// <param name="arguments">JSON arguments object containing operation parameters.</param>
    /// <returns>Result message as a string.</returns>
    /// <exception cref="ArgumentException">Thrown when operation is unknown.</exception>
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
    ///     Adds a hyperlink to a shape or specific text portion.
    /// </summary>
    /// <param name="path">PowerPoint file path.</param>
    /// <param name="outputPath">Output file path.</param>
    /// <param name="arguments">
    ///     JSON arguments containing slideIndex, url or slideTargetIndex, optional text, linkText,
    ///     shapeIndex, position.
    /// </param>
    /// <returns>Success message.</returns>
    /// <exception cref="ArgumentException">
    ///     Thrown when slideIndex or slideTargetIndex is out of range, neither url nor
    ///     slideTargetIndex is provided, or linkText is not found in text.
    /// </exception>
    private Task<string> AddHyperlinkAsync(string path, string outputPath, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var slideIndex = ArgumentHelper.GetInt(arguments, "slideIndex");
            var url = ArgumentHelper.GetStringNullable(arguments, "url");
            var slideTargetIndex = ArgumentHelper.GetIntNullable(arguments, "slideTargetIndex");
            var text = ArgumentHelper.GetStringNullable(arguments, "text");
            var linkText = ArgumentHelper.GetStringNullable(arguments, "linkText");
            var shapeIndex = ArgumentHelper.GetIntNullable(arguments, "shapeIndex");
            var x = ArgumentHelper.GetFloatNullable(arguments, "x");
            var y = ArgumentHelper.GetFloatNullable(arguments, "y");
            var width = ArgumentHelper.GetFloatNullable(arguments, "width");
            var height = ArgumentHelper.GetFloatNullable(arguments, "height");

            using var presentation = new Presentation(path);
            var slide = PowerPointHelper.GetSlide(presentation, slideIndex);
            IAutoShape autoShape;

            if (shapeIndex is >= 0 && shapeIndex.Value < slide.Shapes.Count)
            {
                if (slide.Shapes[shapeIndex.Value] is IAutoShape existingAutoShape)
                    autoShape = existingAutoShape;
                else
                    throw new ArgumentException($"Shape at index {shapeIndex.Value} is not an AutoShape");
            }
            else
            {
                var defaultX = x ?? 50;
                var defaultY = y ?? 50;
                var defaultWidth = width ?? 300;
                var defaultHeight = height ?? 50;
                autoShape = slide.Shapes.AddAutoShape(ShapeType.Rectangle, defaultX, defaultY, defaultWidth,
                    defaultHeight);
            }

            // Create hyperlink object
            IHyperlink hyperlink;
            string linkDescription;
            if (!string.IsNullOrEmpty(url))
            {
                hyperlink = new Hyperlink(url);
                linkDescription = url;
            }
            else if (slideTargetIndex.HasValue)
            {
                PowerPointHelper.ValidateSlideIndex(slideTargetIndex.Value, presentation);
                hyperlink = new Hyperlink(presentation.Slides[slideTargetIndex.Value]);
                linkDescription = $"Slide {slideTargetIndex.Value}";
            }
            else
            {
                throw new ArgumentException("Either url or slideTargetIndex must be provided");
            }

            // Check if we need to apply hyperlink to specific text (Portion-level)
            if (!string.IsNullOrEmpty(linkText) && !string.IsNullOrEmpty(text))
            {
                var linkIndex = text.IndexOf(linkText, StringComparison.Ordinal);
                if (linkIndex < 0)
                    throw new ArgumentException($"linkText '{linkText}' not found in text '{text}'");

                // Clear existing text and create portions
                autoShape.TextFrame.Paragraphs.Clear();
                var paragraph = new Paragraph();

                // Text before the link
                if (linkIndex > 0)
                {
                    var beforePortion = new Portion(text[..linkIndex]);
                    paragraph.Portions.Add(beforePortion);
                }

                // The link text portion
                var linkPortion = new Portion(linkText)
                {
                    PortionFormat = { HyperlinkClick = hyperlink }
                };
                paragraph.Portions.Add(linkPortion);

                // Text after the link
                var afterIndex = linkIndex + linkText.Length;
                if (afterIndex < text.Length)
                {
                    var afterPortion = new Portion(text[afterIndex..]);
                    paragraph.Portions.Add(afterPortion);
                }

                autoShape.TextFrame.Paragraphs.Add(paragraph);
                linkDescription += $" (on text: '{linkText}')";
            }
            else
            {
                // Shape-level hyperlink (original behavior)
                autoShape.HyperlinkClick = hyperlink;

                if (!string.IsNullOrEmpty(text) && autoShape.TextFrame != null)
                    autoShape.TextFrame.Text = text;
            }

            presentation.Save(outputPath, SaveFormat.Pptx);
            return $"Hyperlink added to slide {slideIndex}: {linkDescription}. Output: {outputPath}";
        });
    }

    /// <summary>
    ///     Edits an existing hyperlink.
    /// </summary>
    /// <param name="path">PowerPoint file path.</param>
    /// <param name="outputPath">Output file path.</param>
    /// <param name="arguments">
    ///     JSON arguments containing slideIndex, shapeIndex, optional url, slideTargetIndex,
    ///     removeHyperlink.
    /// </param>
    /// <returns>Success message.</returns>
    /// <exception cref="ArgumentException">
    ///     Thrown when slideIndex, shapeIndex, or slideTargetIndex is out of range, or no
    ///     valid action is specified.
    /// </exception>
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
    ///     Deletes a hyperlink from a shape.
    /// </summary>
    /// <param name="path">PowerPoint file path.</param>
    /// <param name="outputPath">Output file path.</param>
    /// <param name="arguments">JSON arguments containing slideIndex, shapeIndex.</param>
    /// <returns>Success message.</returns>
    /// <exception cref="ArgumentException">Thrown when slideIndex or shapeIndex is out of range.</exception>
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
    ///     Gets all hyperlinks from the presentation.
    /// </summary>
    /// <param name="path">PowerPoint file path.</param>
    /// <param name="arguments">JSON arguments containing optional slideIndex.</param>
    /// <returns>JSON string with all hyperlinks.</returns>
    /// <exception cref="ArgumentException">Thrown when slideIndex is out of range.</exception>
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

    /// <summary>
    ///     Gets hyperlinks from a slide as JSON objects.
    ///     Detects both shape-level and portion-level (text) hyperlinks.
    /// </summary>
    /// <param name="presentation">Presentation to get slide indices from.</param>
    /// <param name="slide">Slide to get hyperlinks from.</param>
    /// <returns>List of hyperlink objects.</returns>
    private static List<object> GetHyperlinksFromSlideAsJson(IPresentation presentation, ISlide slide)
    {
        var hyperlinksList = new List<object>();

        for (var shapeIndex = 0; shapeIndex < slide.Shapes.Count; shapeIndex++)
        {
            if (slide.Shapes[shapeIndex] is not IAutoShape autoShape) continue;

            // Shape-level hyperlinks
            if (autoShape.HyperlinkClick != null)
            {
                var targetSlide = autoShape.HyperlinkClick.TargetSlide;
                var url = autoShape.HyperlinkClick.ExternalUrl
                          ?? (targetSlide != null
                              ? $"Slide {presentation.Slides.IndexOf(targetSlide)}"
                              : "Internal link");

                hyperlinksList.Add(new
                {
                    shapeIndex,
                    level = "shape",
                    triggerType = "click",
                    url
                });
            }

            if (autoShape.HyperlinkMouseOver != null)
            {
                var targetSlide = autoShape.HyperlinkMouseOver.TargetSlide;
                var url = autoShape.HyperlinkMouseOver.ExternalUrl
                          ?? (targetSlide != null
                              ? $"Slide {presentation.Slides.IndexOf(targetSlide)}"
                              : "Internal link");

                hyperlinksList.Add(new
                {
                    shapeIndex,
                    level = "shape",
                    triggerType = "mouseover",
                    url
                });
            }

            // Portion-level (text) hyperlinks
            if (autoShape.TextFrame == null) continue;

            foreach (var paragraph in autoShape.TextFrame.Paragraphs)
            foreach (var portion in paragraph.Portions)
            {
                if (portion.PortionFormat.HyperlinkClick != null)
                {
                    var targetSlide = portion.PortionFormat.HyperlinkClick.TargetSlide;
                    var url = portion.PortionFormat.HyperlinkClick.ExternalUrl
                              ?? (targetSlide != null
                                  ? $"Slide {presentation.Slides.IndexOf(targetSlide)}"
                                  : "Internal link");

                    hyperlinksList.Add(new
                    {
                        shapeIndex,
                        level = "text",
                        triggerType = "click",
                        text = portion.Text,
                        url
                    });
                }

                if (portion.PortionFormat.HyperlinkMouseOver != null)
                {
                    var targetSlide = portion.PortionFormat.HyperlinkMouseOver.TargetSlide;
                    var url = portion.PortionFormat.HyperlinkMouseOver.ExternalUrl
                              ?? (targetSlide != null
                                  ? $"Slide {presentation.Slides.IndexOf(targetSlide)}"
                                  : "Internal link");

                    hyperlinksList.Add(new
                    {
                        shapeIndex,
                        level = "text",
                        triggerType = "mouseover",
                        text = portion.Text,
                        url
                    });
                }
            }
        }

        return hyperlinksList;
    }
}