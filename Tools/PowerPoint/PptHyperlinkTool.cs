using System.ComponentModel;
using System.Text.Json;
using Aspose.Slides;
using AsposeMcpServer.Core.Helpers;
using AsposeMcpServer.Core.Session;
using ModelContextProtocol.Server;

namespace AsposeMcpServer.Tools.PowerPoint;

/// <summary>
///     Unified tool for managing PowerPoint hyperlinks (add, edit, delete, get)
/// </summary>
[McpServerToolType]
public class PptHyperlinkTool
{
    /// <summary>
    ///     Identity accessor for session isolation.
    /// </summary>
    private readonly ISessionIdentityAccessor? _identityAccessor;

    /// <summary>
    ///     Session manager for document lifecycle management.
    /// </summary>
    private readonly DocumentSessionManager? _sessionManager;

    /// <summary>
    ///     Initializes a new instance of the <see cref="PptHyperlinkTool" /> class.
    /// </summary>
    /// <param name="sessionManager">Optional session manager for in-memory document editing.</param>
    /// <param name="identityAccessor">Optional identity accessor for session isolation.</param>
    public PptHyperlinkTool(DocumentSessionManager? sessionManager = null,
        ISessionIdentityAccessor? identityAccessor = null)
    {
        _sessionManager = sessionManager;
        _identityAccessor = identityAccessor;
    }

    /// <summary>
    ///     Executes a PowerPoint hyperlink operation (add, edit, delete, get).
    /// </summary>
    /// <param name="operation">The operation to perform: add, edit, delete, get.</param>
    /// <param name="path">Presentation file path (required if no sessionId).</param>
    /// <param name="sessionId">Session ID for in-memory editing.</param>
    /// <param name="outputPath">Output file path (file mode only).</param>
    /// <param name="slideIndex">Slide index (0-based).</param>
    /// <param name="shapeIndex">Shape index (0-based, required for edit/delete, optional for add).</param>
    /// <param name="text">Display text (required for add).</param>
    /// <param name="linkText">Specific text to apply hyperlink to (optional, for add).</param>
    /// <param name="url">Hyperlink URL (required for add, optional for edit).</param>
    /// <param name="slideTargetIndex">Target slide index for internal link (0-based, optional, for add/edit).</param>
    /// <param name="removeHyperlink">Remove hyperlink (optional, for edit).</param>
    /// <param name="x">X position (optional, for add, default: 50).</param>
    /// <param name="y">Y position (optional, for add, default: 50).</param>
    /// <param name="width">Width (optional, for add, default: 300).</param>
    /// <param name="height">Height (optional, for add, default: 50).</param>
    /// <returns>A message indicating the result of the operation, or JSON data for get operations.</returns>
    /// <exception cref="ArgumentException">Thrown when required parameters are missing or the operation is unknown.</exception>
    [McpServerTool(Name = "ppt_hyperlink")]
    [Description(@"Manage PowerPoint hyperlinks. Supports 4 operations: add, edit, delete, get.

Usage examples:
- Add hyperlink (URL, shape-level): ppt_hyperlink(operation='add', path='presentation.pptx', slideIndex=0, text='Click here', url='https://example.com')
- Add hyperlink (URL, text-level): ppt_hyperlink(operation='add', path='presentation.pptx', slideIndex=0, text='Please click here for more info', linkText='here', url='https://example.com')
- Add hyperlink (internal): ppt_hyperlink(operation='add', path='presentation.pptx', slideIndex=0, text='Go to slide 5', slideTargetIndex=4)
- Edit hyperlink: ppt_hyperlink(operation='edit', path='presentation.pptx', slideIndex=0, shapeIndex=0, url='https://newurl.com')
- Delete hyperlink: ppt_hyperlink(operation='delete', path='presentation.pptx', slideIndex=0, shapeIndex=0)
- Get hyperlinks: ppt_hyperlink(operation='get', path='presentation.pptx', slideIndex=0)")]
    public string Execute(
        [Description("Operation: add, edit, delete, get")]
        string operation,
        [Description("Presentation file path (required if no sessionId)")]
        string? path = null,
        [Description("Session ID for in-memory editing")]
        string? sessionId = null,
        [Description("Output file path (file mode only)")]
        string? outputPath = null,
        [Description("Slide index (0-based)")] int? slideIndex = null,
        [Description("Shape index (0-based, required for edit/delete, optional for add)")]
        int? shapeIndex = null,
        [Description("Display text (required for add)")]
        string? text = null,
        [Description(
            "Specific text to apply hyperlink to (optional, for add). When provided, only this text portion will have the hyperlink.")]
        string? linkText = null,
        [Description("Hyperlink URL (required for add, optional for edit)")]
        string? url = null,
        [Description("Target slide index for internal link (0-based, optional, for add/edit)")]
        int? slideTargetIndex = null,
        [Description("Remove hyperlink (optional, for edit)")]
        bool removeHyperlink = false,
        [Description("X position (optional, for add, default: 50)")]
        float x = 50,
        [Description("Y position (optional, for add, default: 50)")]
        float y = 50,
        [Description("Width (optional, for add, default: 300)")]
        float width = 300,
        [Description("Height (optional, for add, default: 50)")]
        float height = 50)
    {
        using var ctx = DocumentContext<Presentation>.Create(_sessionManager, sessionId, path, _identityAccessor);

        return operation.ToLower() switch
        {
            "add" => AddHyperlink(ctx, outputPath, slideIndex, url, slideTargetIndex, text, linkText, shapeIndex, x, y,
                width, height),
            "edit" => EditHyperlink(ctx, outputPath, slideIndex, shapeIndex, url, slideTargetIndex, removeHyperlink),
            "delete" => DeleteHyperlink(ctx, outputPath, slideIndex, shapeIndex),
            "get" => GetHyperlinks(ctx, slideIndex),
            _ => throw new ArgumentException($"Unknown operation: {operation}")
        };
    }

    /// <summary>
    ///     Adds a hyperlink to a shape or specific text portion.
    /// </summary>
    /// <param name="ctx">The document context.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="slideIndex">The slide index (0-based).</param>
    /// <param name="url">The hyperlink URL.</param>
    /// <param name="slideTargetIndex">The target slide index for internal links.</param>
    /// <param name="text">The display text for the hyperlink.</param>
    /// <param name="linkText">The specific text portion to apply hyperlink to.</param>
    /// <param name="shapeIndex">The shape index to add hyperlink to.</param>
    /// <param name="x">The X position for new shape.</param>
    /// <param name="y">The Y position for new shape.</param>
    /// <param name="width">The width for new shape.</param>
    /// <param name="height">The height for new shape.</param>
    /// <returns>A message indicating the result of the operation.</returns>
    /// <exception cref="ArgumentException">
    ///     Thrown when slideIndex is not provided, or when neither url nor slideTargetIndex is
    ///     provided, or when linkText is not found in text.
    /// </exception>
    private static string AddHyperlink(DocumentContext<Presentation> ctx, string? outputPath,
        int? slideIndex, string? url, int? slideTargetIndex, string? text, string? linkText,
        int? shapeIndex, float x, float y, float width, float height)
    {
        if (!slideIndex.HasValue)
            throw new ArgumentException("slideIndex is required for add operation");

        var presentation = ctx.Document;
        var slide = PowerPointHelper.GetSlide(presentation, slideIndex.Value);
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
            autoShape = slide.Shapes.AddAutoShape(ShapeType.Rectangle, x, y, width, height);
        }

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

        if (!string.IsNullOrEmpty(linkText) && !string.IsNullOrEmpty(text))
        {
            var linkIndex = text.IndexOf(linkText, StringComparison.Ordinal);
            if (linkIndex < 0)
                throw new ArgumentException($"linkText '{linkText}' not found in text '{text}'");

            autoShape.TextFrame.Paragraphs.Clear();
            var paragraph = new Paragraph();

            if (linkIndex > 0)
            {
                var beforePortion = new Portion(text[..linkIndex]);
                paragraph.Portions.Add(beforePortion);
            }

            var linkPortion = new Portion(linkText)
            {
                PortionFormat = { HyperlinkClick = hyperlink }
            };
            paragraph.Portions.Add(linkPortion);

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
            autoShape.HyperlinkClick = hyperlink;

            if (!string.IsNullOrEmpty(text) && autoShape.TextFrame != null)
                autoShape.TextFrame.Text = text;
        }

        ctx.Save(outputPath);

        var result = $"Hyperlink added to slide {slideIndex}: {linkDescription}. ";
        result += ctx.GetOutputMessage(outputPath);
        return result;
    }

    /// <summary>
    ///     Edits an existing hyperlink.
    /// </summary>
    /// <param name="ctx">The document context.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="slideIndex">The slide index (0-based).</param>
    /// <param name="shapeIndex">The shape index (0-based).</param>
    /// <param name="url">The new hyperlink URL.</param>
    /// <param name="slideTargetIndex">The new target slide index for internal links.</param>
    /// <param name="removeHyperlink">Whether to remove the hyperlink.</param>
    /// <returns>A message indicating the result of the operation.</returns>
    /// <exception cref="ArgumentException">
    ///     Thrown when slideIndex or shapeIndex is not provided, or when neither url,
    ///     slideTargetIndex, nor removeHyperlink is provided.
    /// </exception>
    private static string EditHyperlink(DocumentContext<Presentation> ctx, string? outputPath,
        int? slideIndex, int? shapeIndex, string? url, int? slideTargetIndex, bool removeHyperlink)
    {
        if (!slideIndex.HasValue)
            throw new ArgumentException("slideIndex is required for edit operation");
        if (!shapeIndex.HasValue)
            throw new ArgumentException("shapeIndex is required for edit operation");

        var presentation = ctx.Document;
        var slide = PowerPointHelper.GetSlide(presentation, slideIndex.Value);
        var shape = PowerPointHelper.GetShape(slide, shapeIndex.Value);

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

        ctx.Save(outputPath);

        var result = $"Hyperlink updated on slide {slideIndex}, shape {shapeIndex}. ";
        result += ctx.GetOutputMessage(outputPath);
        return result;
    }

    /// <summary>
    ///     Deletes a hyperlink from a shape.
    /// </summary>
    /// <param name="ctx">The document context.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="slideIndex">The slide index (0-based).</param>
    /// <param name="shapeIndex">The shape index (0-based).</param>
    /// <returns>A message indicating the result of the operation.</returns>
    /// <exception cref="ArgumentException">Thrown when slideIndex or shapeIndex is not provided.</exception>
    private static string DeleteHyperlink(DocumentContext<Presentation> ctx, string? outputPath,
        int? slideIndex, int? shapeIndex)
    {
        if (!slideIndex.HasValue)
            throw new ArgumentException("slideIndex is required for delete operation");
        if (!shapeIndex.HasValue)
            throw new ArgumentException("shapeIndex is required for delete operation");

        var presentation = ctx.Document;
        var slide = PowerPointHelper.GetSlide(presentation, slideIndex.Value);
        var shape = PowerPointHelper.GetShape(slide, shapeIndex.Value);
        shape.HyperlinkClick = null;

        if (shape is IAutoShape { TextFrame: not null } autoShape)
            foreach (var paragraph in autoShape.TextFrame.Paragraphs)
            foreach (var portion in paragraph.Portions)
                portion.PortionFormat.HyperlinkClick = null;

        ctx.Save(outputPath);

        var result = $"Hyperlink deleted from slide {slideIndex}, shape {shapeIndex}. ";
        result += ctx.GetOutputMessage(outputPath);
        return result;
    }

    /// <summary>
    ///     Gets all hyperlinks from the presentation.
    /// </summary>
    /// <param name="ctx">The document context.</param>
    /// <param name="slideIndex">The slide index (0-based), or null for all slides.</param>
    /// <returns>A JSON string containing the hyperlinks information.</returns>
    /// <exception cref="ArgumentException">Thrown when slideIndex is out of range.</exception>
    private static string GetHyperlinks(DocumentContext<Presentation> ctx, int? slideIndex)
    {
        var presentation = ctx.Document;

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
            List<object> slidesList = [];
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
    }

    /// <summary>
    ///     Gets hyperlinks from a slide as JSON objects.
    ///     Detects both shape-level and portion-level (text) hyperlinks.
    /// </summary>
    /// <param name="presentation">The presentation object.</param>
    /// <param name="slide">The slide to extract hyperlinks from.</param>
    /// <returns>A list of hyperlink objects containing shape index, level, trigger type, and URL.</returns>
    private static List<object> GetHyperlinksFromSlideAsJson(IPresentation presentation, ISlide slide)
    {
        List<object> hyperlinksList = [];

        for (var shapeIndex = 0; shapeIndex < slide.Shapes.Count; shapeIndex++)
        {
            if (slide.Shapes[shapeIndex] is not IAutoShape autoShape) continue;

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