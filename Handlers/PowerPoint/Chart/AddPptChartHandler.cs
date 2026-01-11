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
        var slideIndex = parameters.GetRequired<int>("slideIndex");
        var chartTypeStr = parameters.GetRequired<string>("chartType");
        var title = parameters.GetOptional<string?>("title");
        var x = parameters.GetOptional("x", 50f);
        var y = parameters.GetOptional("y", 50f);
        var width = parameters.GetOptional("width", 500f);
        var height = parameters.GetOptional("height", 400f);

        var presentation = context.Document;
        var slide = PowerPointHelper.GetSlide(presentation, slideIndex);

        var chartType = PptChartHelper.ParseChartType(chartTypeStr);
        var chart = slide.Shapes.AddChart(chartType, x, y, width, height);

        if (!string.IsNullOrEmpty(title))
            PptChartHelper.SetChartTitle(chart, title);

        MarkModified(context);

        return Success($"Chart '{chartTypeStr}' added to slide {slideIndex}.");
    }
}
