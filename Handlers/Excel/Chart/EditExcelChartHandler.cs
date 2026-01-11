using Aspose.Cells;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Helpers;

namespace AsposeMcpServer.Handlers.Excel.Chart;

/// <summary>
///     Handler for editing Excel chart properties.
/// </summary>
public class EditExcelChartHandler : OperationHandlerBase<Workbook>
{
    /// <inheritdoc />
    public override string Operation => "edit";

    /// <summary>
    ///     Edits chart properties.
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="parameters">
    ///     Optional: sheetIndex, chartIndex, title, dataRange, categoryAxisDataRange, chartType, showLegend, legendPosition
    /// </param>
    /// <returns>Success message with edit details.</returns>
    public override string Execute(OperationContext<Workbook> context, OperationParameters parameters)
    {
        var sheetIndex = parameters.GetOptional("sheetIndex", 0);
        var chartIndex = parameters.GetOptional("chartIndex", 0);
        var title = parameters.GetOptional<string?>("title");
        var dataRange = parameters.GetOptional<string?>("dataRange");
        var categoryAxisDataRange = parameters.GetOptional<string?>("categoryAxisDataRange");
        var chartTypeStr = parameters.GetOptional<string?>("chartType");
        var showLegend = parameters.GetOptional<bool?>("showLegend");
        var legendPosition = parameters.GetOptional<string?>("legendPosition");

        var workbook = context.Document;
        var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);
        var chart = ExcelChartHelper.GetChart(worksheet, chartIndex);

        List<string> changes = [];

        if (!string.IsNullOrEmpty(title))
        {
            chart.Title.Text = title;
            changes.Add($"Title: {title}");
        }

        if (!string.IsNullOrEmpty(dataRange))
        {
            ExcelChartHelper.AddDataSeries(chart, dataRange);
            ExcelChartHelper.SetCategoryData(chart, categoryAxisDataRange ?? "");

            var rangeInfo = dataRange;
            if (!string.IsNullOrEmpty(categoryAxisDataRange))
                rangeInfo += $", X-axis: {categoryAxisDataRange}";
            changes.Add($"Data range: {rangeInfo}");
        }

        if (!string.IsNullOrEmpty(chartTypeStr))
        {
            chart.Type = ExcelChartHelper.ParseChartType(chartTypeStr, chart.Type);
            changes.Add($"Chart type: {chartTypeStr}");
        }

        if (showLegend.HasValue)
        {
            chart.ShowLegend = showLegend.Value;
            changes.Add($"Legend: {(showLegend.Value ? "show" : "hide")}");
        }

        if (!string.IsNullOrEmpty(legendPosition) && chart.Legend != null)
        {
            chart.Legend.Position = ExcelChartHelper.ParseLegendPosition(legendPosition, chart.Legend.Position);
            changes.Add($"Legend position: {legendPosition}");
        }

        if (changes.Count > 0)
            MarkModified(context);

        return changes.Count > 0
            ? Success($"Chart #{chartIndex} edited: {string.Join(", ", changes)}")
            : Success($"Chart #{chartIndex} no changes");
    }
}
