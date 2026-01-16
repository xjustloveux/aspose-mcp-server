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
        var setParams = ExtractSetPropertiesParameters(parameters);

        var workbook = context.Document;
        var worksheet = ExcelHelper.GetWorksheet(workbook, setParams.SheetIndex);
        var chart = ExcelChartHelper.GetChart(worksheet, setParams.ChartIndex);

        List<string> changes = [];

        if (setParams.RemoveTitle)
        {
            chart.Title.Text = "";
            changes.Add("Title removed");
        }
        else if (!string.IsNullOrEmpty(setParams.Title))
        {
            chart.Title.Text = setParams.Title;
            changes.Add($"Title: {setParams.Title}");
        }

        if (setParams.LegendVisible.HasValue)
        {
            chart.ShowLegend = setParams.LegendVisible.Value;
            changes.Add($"Legend: {(setParams.LegendVisible.Value ? "show" : "hide")}");
        }

        if (!string.IsNullOrEmpty(setParams.LegendPosition) && chart.Legend != null)
        {
            chart.Legend.Position =
                ExcelChartHelper.ParseLegendPosition(setParams.LegendPosition, chart.Legend.Position);
            changes.Add($"Legend position: {setParams.LegendPosition}");
        }

        if (changes.Count > 0)
            MarkModified(context);

        return changes.Count > 0
            ? Success($"Chart #{setParams.ChartIndex} properties updated: {string.Join(", ", changes)}")
            : Success($"Chart #{setParams.ChartIndex} no changes");
    }

    /// <summary>
    ///     Extracts set properties parameters from operation parameters.
    /// </summary>
    /// <param name="parameters">The operation parameters.</param>
    /// <returns>The extracted set properties parameters.</returns>
    private static SetPropertiesParameters ExtractSetPropertiesParameters(OperationParameters parameters)
    {
        return new SetPropertiesParameters(
            parameters.GetOptional("sheetIndex", 0),
            parameters.GetOptional("chartIndex", 0),
            parameters.GetOptional<string?>("title"),
            parameters.GetOptional("removeTitle", false),
            parameters.GetOptional<bool?>("legendVisible"),
            parameters.GetOptional<string?>("legendPosition")
        );
    }

    /// <summary>
    ///     Record to hold set properties parameters.
    /// </summary>
    private record SetPropertiesParameters(
        int SheetIndex,
        int ChartIndex,
        string? Title,
        bool RemoveTitle,
        bool? LegendVisible,
        string? LegendPosition);
}
