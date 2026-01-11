using Aspose.Cells;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Helpers;

namespace AsposeMcpServer.Handlers.Excel.Chart;

/// <summary>
///     Handler for setting Excel chart properties (title, legend).
/// </summary>
public class SetExcelChartPropertiesHandler : OperationHandlerBase<Workbook>
{
    /// <inheritdoc />
    public override string Operation => "set_properties";

    /// <summary>
    ///     Sets chart properties (title, legend).
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="parameters">
    ///     Optional: sheetIndex, chartIndex, title, removeTitle, legendVisible, legendPosition
    /// </param>
    /// <returns>Success message with property update details.</returns>
    public override string Execute(OperationContext<Workbook> context, OperationParameters parameters)
    {
        var sheetIndex = parameters.GetOptional("sheetIndex", 0);
        var chartIndex = parameters.GetOptional("chartIndex", 0);
        var title = parameters.GetOptional<string?>("title");
        var removeTitle = parameters.GetOptional("removeTitle", false);
        var legendVisible = parameters.GetOptional<bool?>("legendVisible");
        var legendPosition = parameters.GetOptional<string?>("legendPosition");

        var workbook = context.Document;
        var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);
        var chart = ExcelChartHelper.GetChart(worksheet, chartIndex);

        List<string> changes = [];

        if (removeTitle)
        {
            chart.Title.Text = "";
            changes.Add("Title removed");
        }
        else if (!string.IsNullOrEmpty(title))
        {
            chart.Title.Text = title;
            changes.Add($"Title: {title}");
        }

        if (legendVisible.HasValue)
        {
            chart.ShowLegend = legendVisible.Value;
            changes.Add($"Legend: {(legendVisible.Value ? "show" : "hide")}");
        }

        if (!string.IsNullOrEmpty(legendPosition) && chart.Legend != null)
        {
            chart.Legend.Position = ExcelChartHelper.ParseLegendPosition(legendPosition, chart.Legend.Position);
            changes.Add($"Legend position: {legendPosition}");
        }

        if (changes.Count > 0)
            MarkModified(context);

        return changes.Count > 0
            ? Success($"Chart #{chartIndex} properties updated: {string.Join(", ", changes)}")
            : Success($"Chart #{chartIndex} no changes");
    }
}
