using System.ComponentModel;
using System.Text.Json;
using Aspose.Slides;
using AsposeMcpServer.Core.Helpers;
using AsposeMcpServer.Core.Session;
using ModelContextProtocol.Server;

namespace AsposeMcpServer.Tools.PowerPoint;

/// <summary>
///     Unified tool for PowerPoint data operations (get statistics, get content, get slide details)
///     Merges: PptGetStatisticsTool, PptGetContentTool, PptGetSlideDetailsTool
/// </summary>
[McpServerToolType]
public class PptDataOperationsTool
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
    ///     Initializes a new instance of the <see cref="PptDataOperationsTool" /> class.
    /// </summary>
    /// <param name="sessionManager">Optional session manager for in-memory document editing.</param>
    /// <param name="identityAccessor">Optional identity accessor for session isolation.</param>
    public PptDataOperationsTool(DocumentSessionManager? sessionManager = null,
        ISessionIdentityAccessor? identityAccessor = null)
    {
        _sessionManager = sessionManager;
        _identityAccessor = identityAccessor;
    }

    /// <summary>
    ///     Executes a PowerPoint data operation (get_statistics, get_content, get_slide_details).
    /// </summary>
    /// <param name="operation">The operation to perform: get_statistics, get_content, get_slide_details.</param>
    /// <param name="path">Presentation file path (required if no sessionId).</param>
    /// <param name="sessionId">Session ID for in-memory editing.</param>
    /// <param name="slideIndex">Slide index (0-based, required for get_slide_details).</param>
    /// <param name="includeThumbnail">Include Base64 encoded thumbnail image (optional for get_slide_details, default false).</param>
    /// <returns>A JSON string containing the requested data (statistics, content, or slide details).</returns>
    /// <exception cref="ArgumentException">Thrown when required parameters are missing or the operation is unknown.</exception>
    [McpServerTool(Name = "ppt_data_operations")]
    [Description(@"PowerPoint data operations. Supports 3 operations: get_statistics, get_content, get_slide_details.

Usage examples:
- Get statistics: ppt_data_operations(operation='get_statistics', path='presentation.pptx')
- Get content: ppt_data_operations(operation='get_content', path='presentation.pptx')
- Get slide details: ppt_data_operations(operation='get_slide_details', path='presentation.pptx', slideIndex=0)
- Get slide details with thumbnail: ppt_data_operations(operation='get_slide_details', path='presentation.pptx', slideIndex=0, includeThumbnail=true)")]
    public string Execute(
        [Description("Operation: get_statistics, get_content, get_slide_details")]
        string operation,
        [Description("Presentation file path (required if no sessionId)")]
        string? path = null,
        [Description("Session ID for in-memory editing")]
        string? sessionId = null,
        [Description("Slide index (0-based, required for get_slide_details)")]
        int? slideIndex = null,
        [Description("Include Base64 encoded thumbnail image (optional for get_slide_details, default false)")]
        bool includeThumbnail = false)
    {
        using var ctx = DocumentContext<Presentation>.Create(_sessionManager, sessionId, path, _identityAccessor);

        return operation.ToLower() switch
        {
            "get_statistics" => GetStatistics(ctx),
            "get_content" => GetContent(ctx),
            "get_slide_details" => GetSlideDetails(ctx, slideIndex, includeThumbnail),
            _ => throw new ArgumentException($"Unknown operation: {operation}")
        };
    }

    /// <summary>
    ///     Gets presentation statistics.
    /// </summary>
    /// <param name="ctx">The document context containing the presentation.</param>
    /// <returns>A JSON string containing presentation statistics including slide count, shape count, and media counts.</returns>
    private static string GetStatistics(DocumentContext<Presentation> ctx)
    {
        var presentation = ctx.Document;

        var totalShapes = 0;
        var totalTextCharacters = 0;
        var totalImages = 0;
        var totalTables = 0;
        var totalCharts = 0;
        var totalSmartArt = 0;
        var totalAudio = 0;
        var totalVideo = 0;
        var totalAnimations = 0;
        var totalHyperlinks = 0;
        var totalHiddenSlides = 0;

        foreach (var slide in presentation.Slides)
        {
            if (slide.Hidden) totalHiddenSlides++;
            totalShapes += slide.Shapes.Count;
            totalAnimations += slide.Timeline.MainSequence.Count;

            foreach (var shape in slide.Shapes)
            {
                totalTextCharacters += PowerPointHelper.CountTextCharacters(shape);
                PowerPointHelper.CountShapeTypes(shape, ref totalImages, ref totalTables, ref totalCharts,
                    ref totalSmartArt, ref totalAudio, ref totalVideo);

                if (shape.HyperlinkClick != null) totalHyperlinks++;
            }
        }

        var result = new
        {
            totalSlides = presentation.Slides.Count,
            totalHiddenSlides,
            totalLayouts = presentation.LayoutSlides.Count,
            totalMasters = presentation.Masters.Count,
            slideSize = new
            {
                width = presentation.SlideSize.Size.Width,
                height = presentation.SlideSize.Size.Height
            },
            totalShapes,
            totalTextCharacters,
            totalImages,
            totalTables,
            totalCharts,
            totalSmartArt,
            totalAudio,
            totalVideo,
            totalAnimations,
            totalHyperlinks
        };

        return JsonSerializer.Serialize(result, new JsonSerializerOptions { WriteIndented = true });
    }

    /// <summary>
    ///     Gets presentation content including text from all shape types.
    /// </summary>
    /// <param name="ctx">The document context containing the presentation.</param>
    /// <returns>A JSON string containing text content extracted from all slides.</returns>
    private static string GetContent(DocumentContext<Presentation> ctx)
    {
        var presentation = ctx.Document;
        List<object> slides = [];

        var slideIndex = 0;
        foreach (var slide in presentation.Slides)
        {
            List<string> textContent = [];
            foreach (var shape in slide.Shapes)
                PowerPointHelper.ExtractTextFromShape(shape, textContent);

            slides.Add(new
            {
                index = slideIndex,
                hidden = slide.Hidden,
                textContent
            });
            slideIndex++;
        }

        var result = new
        {
            totalSlides = presentation.Slides.Count,
            slides
        };

        return JsonSerializer.Serialize(result, new JsonSerializerOptions { WriteIndented = true });
    }

    /// <summary>
    ///     Gets detailed information about a slide.
    /// </summary>
    /// <param name="ctx">The document context containing the presentation.</param>
    /// <param name="slideIndex">The zero-based index of the slide to get details for.</param>
    /// <param name="includeThumbnail">Whether to include a Base64 encoded thumbnail image.</param>
    /// <returns>A JSON string containing detailed slide information including layout, transitions, and animations.</returns>
    /// <exception cref="ArgumentException">Thrown when slideIndex is not provided.</exception>
    private static string GetSlideDetails(DocumentContext<Presentation> ctx, int? slideIndex, bool includeThumbnail)
    {
        if (!slideIndex.HasValue)
            throw new ArgumentException("slideIndex is required for get_slide_details operation");

        var presentation = ctx.Document;
        var slide = PowerPointHelper.GetSlide(presentation, slideIndex.Value);

        var transition = slide.SlideShowTransition;
        object? transitionInfo = transition != null
            ? new
            {
                type = transition.Type.ToString(),
                speed = transition.Speed.ToString(),
                advanceOnClick = transition.AdvanceOnClick,
                advanceAfterTimeMs = transition.AdvanceAfterTime
            }
            : null;

        var animations = slide.Timeline.MainSequence;
        List<object> animationsList = [];
        for (var i = 0; i < animations.Count; i++)
        {
            var anim = animations[i];
            animationsList.Add(new
            {
                index = i,
                type = anim.Type.ToString(),
                targetShape = anim.TargetShape?.GetType().Name
            });
        }

        var background = slide.Background;
        object? backgroundInfo = background != null
            ? new { fillType = background.FillFormat.FillType.ToString() }
            : null;

        var notesSlide = slide.NotesSlideManager.NotesSlide;
        var notesText = notesSlide?.NotesTextFrame?.Text;

        string? thumbnailBase64 = null;
        if (includeThumbnail) thumbnailBase64 = PowerPointHelper.GenerateThumbnail(slide);

        var result = new
        {
            slideIndex,
            hidden = slide.Hidden,
            layout = slide.LayoutSlide?.Name,
            slideSize = new
            {
                width = presentation.SlideSize.Size.Width,
                height = presentation.SlideSize.Size.Height
            },
            shapesCount = slide.Shapes.Count,
            transition = transitionInfo,
            animationsCount = animations.Count,
            animations = animationsList,
            background = backgroundInfo,
            notes = string.IsNullOrWhiteSpace(notesText) ? null : notesText,
            thumbnail = thumbnailBase64
        };

        return JsonSerializer.Serialize(result, new JsonSerializerOptions { WriteIndented = true });
    }
}