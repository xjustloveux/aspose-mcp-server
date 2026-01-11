using Aspose.Cells;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Helpers;

namespace AsposeMcpServer.Handlers.Excel.Chart;

/// <summary>
///     Handler for deleting charts from Excel worksheets.
/// </summary>
public class DeleteExcelChartHandler : OperationHandlerBase<Workbook>
{
    /// <inheritdoc />
    public override string Operation => "delete";

    /// <summary>
    ///     Deletes a chart from the worksheet.
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="parameters">
    ///     Optional: sheetIndex, chartIndex
    /// </param>
    /// <returns>Success message with deletion details.</returns>
    public override string Execute(OperationContext<Workbook> context, OperationParameters parameters)
    {
        var sheetIndex = parameters.GetOptional("sheetIndex", 0);
        var chartIndex = parameters.GetOptional("chartIndex", 0);

        var workbook = context.Document;
        var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);
        var chart = ExcelChartHelper.GetChart(worksheet, chartIndex);

        var chartName = chart.Name ?? $"Chart {chartIndex}";
        worksheet.Charts.RemoveAt(chartIndex);

        MarkModified(context);

        return Success($"Chart #{chartIndex} ({chartName}) deleted, {worksheet.Charts.Count} remaining");
    }
}
