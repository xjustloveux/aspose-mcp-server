using Aspose.Slides;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Helpers;

namespace AsposeMcpServer.Handlers.PowerPoint.Chart;

/// <summary>
///     Handler for editing chart properties in PowerPoint presentations.
/// </summary>
public class EditPptChartHandler : OperationHandlerBase<Presentation>
{
    /// <inheritdoc />
    public override string Operation => "edit";

    /// <summary>
    ///     Edits chart properties.
    /// </summary>
    /// <param name="context">The presentation context.</param>
    /// <param name="parameters">
    ///     Required: slideIndex, shapeIndex
    ///     Optional: title, chartType
    /// </param>
    /// <returns>Success message with update details.</returns>
    public override string Execute(OperationContext<Presentation> context, OperationParameters parameters)
    {
        var slideIndex = parameters.GetRequired<int>("slideIndex");
        var chartIndex = parameters.GetRequired<int>("shapeIndex");
        var title = parameters.GetOptional<string?>("title");
        var chartTypeStr = parameters.GetOptional<string?>("chartType");

        var presentation = context.Document;
        var slide = PowerPointHelper.GetSlide(presentation, slideIndex);
        var chart = PptChartHelper.GetChartByIndex(slide, chartIndex, slideIndex);

        if (!string.IsNullOrEmpty(title))
            try
            {
                PptChartHelper.SetChartTitle(chart, title);
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException($"Failed to set chart title: {ex.Message}", ex);
            }

        if (!string.IsNullOrEmpty(chartTypeStr))
            try
            {
                chart.Type = PptChartHelper.ParseChartType(chartTypeStr, chart.Type);
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException($"Failed to change chart type: {ex.Message}", ex);
            }

        MarkModified(context);

        return Success($"Chart {chartIndex} updated on slide {slideIndex}.");
    }
}
