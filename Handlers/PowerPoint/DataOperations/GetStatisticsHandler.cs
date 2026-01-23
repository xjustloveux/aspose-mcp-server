using Aspose.Slides;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Helpers.PowerPoint;
using AsposeMcpServer.Results.PowerPoint.DataOperations;

namespace AsposeMcpServer.Handlers.PowerPoint.DataOperations;

/// <summary>
///     Handler for getting presentation statistics.
/// </summary>
[ResultType(typeof(GetStatisticsResult))]
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
    public override object Execute(OperationContext<Presentation> context, OperationParameters parameters)
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

        var result = new GetStatisticsResult
        {
            TotalSlides = presentation.Slides.Count,
            TotalHiddenSlides = totalHiddenSlides,
            TotalLayouts = presentation.LayoutSlides.Count,
            TotalMasters = presentation.Masters.Count,
            SlideSize = new GetStatisticsSizeInfo
            {
                Width = presentation.SlideSize.Size.Width,
                Height = presentation.SlideSize.Size.Height
            },
            TotalShapes = totalShapes,
            TotalTextCharacters = totalTextCharacters,
            TotalImages = totalImages,
            TotalTables = totalTables,
            TotalCharts = totalCharts,
            TotalSmartArt = totalSmartArt,
            TotalAudio = totalAudio,
            TotalVideo = totalVideo,
            TotalAnimations = totalAnimations,
            TotalHyperlinks = totalHyperlinks
        };

        return result;
    }
}
