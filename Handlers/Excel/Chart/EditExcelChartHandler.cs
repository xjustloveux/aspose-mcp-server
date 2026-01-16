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
        var editParams = ExtractEditParameters(parameters);

        var workbook = context.Document;
        var worksheet = ExcelHelper.GetWorksheet(workbook, editParams.SheetIndex);
        var chart = ExcelChartHelper.GetChart(worksheet, editParams.ChartIndex);

        List<string> changes = [];

        if (!string.IsNullOrEmpty(editParams.Title))
        {
            chart.Title.Text = editParams.Title;
            changes.Add($"Title: {editParams.Title}");
        }

        if (!string.IsNullOrEmpty(editParams.DataRange))
        {
            ExcelChartHelper.AddDataSeries(chart, editParams.DataRange);
            ExcelChartHelper.SetCategoryData(chart, editParams.CategoryAxisDataRange ?? "");

            var rangeInfo = editParams.DataRange;
            if (!string.IsNullOrEmpty(editParams.CategoryAxisDataRange))
                rangeInfo += $", X-axis: {editParams.CategoryAxisDataRange}";
            changes.Add($"Data range: {rangeInfo}");
        }

        if (!string.IsNullOrEmpty(editParams.ChartType))
        {
            chart.Type = ExcelChartHelper.ParseChartType(editParams.ChartType, chart.Type);
            changes.Add($"Chart type: {editParams.ChartType}");
        }

        if (editParams.ShowLegend.HasValue)
        {
            chart.ShowLegend = editParams.ShowLegend.Value;
            changes.Add($"Legend: {(editParams.ShowLegend.Value ? "show" : "hide")}");
        }

        if (!string.IsNullOrEmpty(editParams.LegendPosition) && chart.Legend != null)
        {
            chart.Legend.Position =
                ExcelChartHelper.ParseLegendPosition(editParams.LegendPosition, chart.Legend.Position);
            changes.Add($"Legend position: {editParams.LegendPosition}");
        }

        if (changes.Count > 0)
            MarkModified(context);

        return changes.Count > 0
            ? Success($"Chart #{editParams.ChartIndex} edited: {string.Join(", ", changes)}")
            : Success($"Chart #{editParams.ChartIndex} no changes");
    }

    /// <summary>
    ///     Extracts edit parameters from operation parameters.
    /// </summary>
    /// <param name="parameters">The operation parameters.</param>
    /// <returns>The extracted edit parameters.</returns>
    private static EditParameters ExtractEditParameters(OperationParameters parameters)
    {
        return new EditParameters(
            parameters.GetOptional("sheetIndex", 0),
            parameters.GetOptional("chartIndex", 0),
            parameters.GetOptional<string?>("title"),
            parameters.GetOptional<string?>("dataRange"),
            parameters.GetOptional<string?>("categoryAxisDataRange"),
            parameters.GetOptional<string?>("chartType"),
            parameters.GetOptional<bool?>("showLegend"),
            parameters.GetOptional<string?>("legendPosition")
        );
    }

    /// <summary>
    ///     Record to hold edit chart parameters.
    /// </summary>
    private sealed record EditParameters(
        int SheetIndex,
        int ChartIndex,
        string? Title,
        string? DataRange,
        string? CategoryAxisDataRange,
        string? ChartType,
        bool? ShowLegend,
        string? LegendPosition);
}
