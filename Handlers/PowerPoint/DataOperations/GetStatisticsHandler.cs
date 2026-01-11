using Aspose.Slides;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Helpers;

namespace AsposeMcpServer.Handlers.PowerPoint.DataOperations;

/// <summary>
///     Handler for getting presentation statistics.
/// </summary>
public class GetStatisticsHandler : OperationHandlerBase<Presentation>
{
    /// <inheritdoc />
    public override string Operation => "get_statistics";

    /// <summary>
    ///     Gets presentation statistics including slide count, shape count, and media counts.
    /// </summary>
    /// <param name="context">The presentation context.</param>
    /// <param name="parameters">No parameters required.</param>
    /// <returns>JSON string containing the presentation statistics.</returns>
    public override string Execute(OperationContext<Presentation> context, OperationParameters parameters)
    {
        var presentation = context.Document;

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

        return JsonResult(result);
    }
}
