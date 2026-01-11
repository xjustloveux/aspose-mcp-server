using Aspose.Cells;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Helpers;

namespace AsposeMcpServer.Handlers.Excel.Chart;

/// <summary>
///     Handler for updating Excel chart data range.
/// </summary>
public class UpdateExcelChartDataHandler : OperationHandlerBase<Workbook>
{
    /// <inheritdoc />
    public override string Operation => "update_data";

    /// <summary>
    ///     Updates chart data range.
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="parameters">
    ///     Required: dataRange
    ///     Optional: sheetIndex, chartIndex, categoryAxisDataRange
    /// </param>
    /// <returns>Success message with update details.</returns>
    public override string Execute(OperationContext<Workbook> context, OperationParameters parameters)
    {
        var sheetIndex = parameters.GetOptional("sheetIndex", 0);
        var chartIndex = parameters.GetOptional("chartIndex", 0);
        var dataRange = parameters.GetOptional<string?>("dataRange");
        var categoryAxisDataRange = parameters.GetOptional<string?>("categoryAxisDataRange");

        if (string.IsNullOrEmpty(dataRange))
            throw new ArgumentException("dataRange is required for update_data operation");

        var workbook = context.Document;
        var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);
        var chart = ExcelChartHelper.GetChart(worksheet, chartIndex);

        ExcelChartHelper.AddDataSeries(chart, dataRange);
        ExcelChartHelper.SetCategoryData(chart, categoryAxisDataRange ?? "");

        MarkModified(context);

        var result = $"Chart #{chartIndex} data updated to: {dataRange}";
        if (!string.IsNullOrEmpty(categoryAxisDataRange))
            result += $", X-axis: {categoryAxisDataRange}";
        return Success(result);
    }
}
