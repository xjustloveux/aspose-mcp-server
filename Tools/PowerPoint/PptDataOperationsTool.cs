using System.Text.Json;
using System.Text.Json.Nodes;
using Aspose.Slides;
using AsposeMcpServer.Core;

namespace AsposeMcpServer.Tools.PowerPoint;

/// <summary>
///     Unified tool for PowerPoint data operations (get statistics, get content, get slide details)
///     Merges: PptGetStatisticsTool, PptGetContentTool, PptGetSlideDetailsTool
/// </summary>
public class PptDataOperationsTool : IAsposeTool
{
    public string Description =>
        @"PowerPoint data operations. Supports 3 operations: get_statistics, get_content, get_slide_details.

Usage examples:
- Get statistics: ppt_data_operations(operation='get_statistics', path='presentation.pptx')
- Get content: ppt_data_operations(operation='get_content', path='presentation.pptx')
- Get slide details: ppt_data_operations(operation='get_slide_details', path='presentation.pptx', slideIndex=0)
- Get slide details with thumbnail: ppt_data_operations(operation='get_slide_details', path='presentation.pptx', slideIndex=0, includeThumbnail=true)";

    public object InputSchema => new
    {
        type = "object",
        properties = new
        {
            operation = new
            {
                type = "string",
                description = @"Operation to perform.
- 'get_statistics': Get presentation statistics (required params: path)
- 'get_content': Get presentation content (required params: path)
- 'get_slide_details': Get slide details (required params: path, slideIndex)",
                @enum = new[] { "get_statistics", "get_content", "get_slide_details" }
            },
            path = new
            {
                type = "string",
                description = "Presentation file path (required for all operations)"
            },
            slideIndex = new
            {
                type = "number",
                description = "Slide index (0-based, required for get_slide_details)"
            },
            includeThumbnail = new
            {
                type = "boolean",
                description = "Include Base64 encoded thumbnail image (optional for get_slide_details, default false)"
            }
        },
        required = new[] { "operation", "path" }
    };

    /// <inheritdoc />
    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var operation = ArgumentHelper.GetString(arguments, "operation");
        var path = ArgumentHelper.GetAndValidatePath(arguments);

        return operation.ToLower() switch
        {
            "get_statistics" => await GetStatisticsAsync(arguments, path),
            "get_content" => await GetContentAsync(arguments, path),
            "get_slide_details" => await GetSlideDetailsAsync(arguments, path),
            _ => throw new ArgumentException($"Unknown operation: {operation}")
        };
    }

    /// <summary>
    ///     Gets presentation statistics.
    /// </summary>
    /// <param name="_">Unused parameter.</param>
    /// <param name="path">PowerPoint file path.</param>
    /// <returns>JSON string with presentation statistics.</returns>
    private Task<string> GetStatisticsAsync(JsonObject? _, string path)
    {
        return Task.Run(() =>
        {
            using var presentation = new Presentation(path);

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
        });
    }

    /// <summary>
    ///     Gets presentation content including text from all shape types.
    /// </summary>
    /// <param name="_">Unused parameter.</param>
    /// <param name="path">PowerPoint file path.</param>
    /// <returns>JSON string with presentation content.</returns>
    private Task<string> GetContentAsync(JsonObject? _, string path)
    {
        return Task.Run(() =>
        {
            using var presentation = new Presentation(path);
            var slides = new List<object>();

            var slideIndex = 0;
            foreach (var slide in presentation.Slides)
            {
                var textContent = new List<string>();
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
        });
    }

    /// <summary>
    ///     Gets detailed information about a slide.
    /// </summary>
    /// <param name="arguments">JSON arguments containing slideIndex and optional includeThumbnail.</param>
    /// <param name="path">PowerPoint file path.</param>
    /// <returns>JSON string with slide details.</returns>
    /// <exception cref="ArgumentException">Thrown when slideIndex is out of range.</exception>
    private Task<string> GetSlideDetailsAsync(JsonObject? arguments, string path)
    {
        return Task.Run(() =>
        {
            var slideIndex = ArgumentHelper.GetInt(arguments, "slideIndex");
            var includeThumbnail = ArgumentHelper.GetBool(arguments, "includeThumbnail", false);

            using var presentation = new Presentation(path);
            var slide = PowerPointHelper.GetSlide(presentation, slideIndex);

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
            var animationsList = new List<object>();
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

            // Generate thumbnail if requested
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
        });
    }
}