using Aspose.Slides;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Helpers;

namespace AsposeMcpServer.Handlers.PowerPoint.Chart;

/// <summary>
///     Handler for adding charts to PowerPoint presentations.
/// </summary>
public class AddPptChartHandler : OperationHandlerBase<Presentation>
{
    /// <inheritdoc />
    public override string Operation => "add";

    /// <summary>
    ///     Adds a chart to a slide.
    /// </summary>
    /// <param name="context">The presentation context.</param>
    /// <param name="parameters">
    ///     Required: slideIndex, chartType
    ///     Optional: title, x, y, width, height
    /// </param>
    /// <returns>Success message with chart details.</returns>
    public override string Execute(OperationContext<Presentation> context, OperationParameters parameters)
    {
        var p = ExtractAddChartParameters(parameters);

        var presentation = context.Document;
        var slide = PowerPointHelper.GetSlide(presentation, p.SlideIndex);

        var chartType = PptChartHelper.ParseChartType(p.ChartType);
        var chart = slide.Shapes.AddChart(chartType, p.X, p.Y, p.Width, p.Height);

        if (!string.IsNullOrEmpty(p.Title))
            PptChartHelper.SetChartTitle(chart, p.Title);

        MarkModified(context);

        return Success($"Chart '{p.ChartType}' added to slide {p.SlideIndex}.");
    }

    /// <summary>
    ///     Extracts add chart parameters from operation parameters.
    /// </summary>
    /// <param name="parameters">The operation parameters.</param>
    /// <returns>The extracted add chart parameters.</returns>
    private static AddChartParameters ExtractAddChartParameters(OperationParameters parameters)
    {
        return new AddChartParameters(
            parameters.GetRequired<int>("slideIndex"),
            parameters.GetRequired<string>("chartType"),
            parameters.GetOptional<string?>("title"),
            parameters.GetOptional("x", 50f),
            parameters.GetOptional("y", 50f),
            parameters.GetOptional("width", 500f),
            parameters.GetOptional("height", 400f));
    }

    /// <summary>
    ///     Record for holding add chart parameters.
    /// </summary>
    /// <param name="SlideIndex">The slide index.</param>
    /// <param name="ChartType">The chart type.</param>
    /// <param name="Title">The optional chart title.</param>
    /// <param name="X">The X coordinate.</param>
    /// <param name="Y">The Y coordinate.</param>
    /// <param name="Width">The chart width.</param>
    /// <param name="Height">The chart height.</param>
    private record AddChartParameters(
        int SlideIndex,
        string ChartType,
        string? Title,
        float X,
        float Y,
        float Width,
        float Height);
}
