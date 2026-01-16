using Aspose.Slides;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Helpers;

namespace AsposeMcpServer.Handlers.PowerPoint.Chart;

/// <summary>
///     Handler for deleting charts from PowerPoint presentations.
/// </summary>
public class DeletePptChartHandler : OperationHandlerBase<Presentation>
{
    /// <inheritdoc />
    public override string Operation => "delete";

    /// <summary>
    ///     Deletes a chart from a slide.
    /// </summary>
    /// <param name="context">The presentation context.</param>
    /// <param name="parameters">
    ///     Required: slideIndex, shapeIndex
    /// </param>
    /// <returns>Success message with deletion details.</returns>
    public override string Execute(OperationContext<Presentation> context, OperationParameters parameters)
    {
        var p = ExtractDeleteChartParameters(parameters);

        var presentation = context.Document;
        var slide = PowerPointHelper.GetSlide(presentation, p.SlideIndex);
        var chart = PptChartHelper.GetChartByIndex(slide, p.ChartIndex, p.SlideIndex);

        slide.Shapes.Remove(chart);

        MarkModified(context);

        return Success($"Chart {p.ChartIndex} deleted from slide {p.SlideIndex}.");
    }

    /// <summary>
    ///     Extracts delete chart parameters from operation parameters.
    /// </summary>
    /// <param name="parameters">The operation parameters.</param>
    /// <returns>The extracted delete chart parameters.</returns>
    private static DeleteChartParameters ExtractDeleteChartParameters(OperationParameters parameters)
    {
        return new DeleteChartParameters(
            parameters.GetRequired<int>("slideIndex"),
            parameters.GetRequired<int>("shapeIndex"));
    }

    /// <summary>
    ///     Record for holding delete chart parameters.
    /// </summary>
    /// <param name="SlideIndex">The slide index.</param>
    /// <param name="ChartIndex">The chart shape index.</param>
    private record DeleteChartParameters(int SlideIndex, int ChartIndex);
}
