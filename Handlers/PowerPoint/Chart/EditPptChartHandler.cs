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
        var p = ExtractEditChartParameters(parameters);

        var presentation = context.Document;
        var slide = PowerPointHelper.GetSlide(presentation, p.SlideIndex);
        var chart = PptChartHelper.GetChartByIndex(slide, p.ChartIndex, p.SlideIndex);

        if (!string.IsNullOrEmpty(p.Title))
            try
            {
                PptChartHelper.SetChartTitle(chart, p.Title);
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException($"Failed to set chart title: {ex.Message}", ex);
            }

        if (!string.IsNullOrEmpty(p.ChartType))
            try
            {
                chart.Type = PptChartHelper.ParseChartType(p.ChartType, chart.Type);
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException($"Failed to change chart type: {ex.Message}", ex);
            }

        MarkModified(context);

        return Success($"Chart {p.ChartIndex} updated on slide {p.SlideIndex}.");
    }

    /// <summary>
    ///     Extracts edit chart parameters from operation parameters.
    /// </summary>
    /// <param name="parameters">The operation parameters.</param>
    /// <returns>The extracted edit chart parameters.</returns>
    private static EditChartParameters ExtractEditChartParameters(OperationParameters parameters)
    {
        return new EditChartParameters(
            parameters.GetRequired<int>("slideIndex"),
            parameters.GetRequired<int>("shapeIndex"),
            parameters.GetOptional<string?>("title"),
            parameters.GetOptional<string?>("chartType"));
    }

    /// <summary>
    ///     Record for holding edit chart parameters.
    /// </summary>
    /// <param name="SlideIndex">The slide index.</param>
    /// <param name="ChartIndex">The chart shape index.</param>
    /// <param name="Title">The optional chart title.</param>
    /// <param name="ChartType">The optional chart type.</param>
    private record EditChartParameters(int SlideIndex, int ChartIndex, string? Title, string? ChartType);
}
