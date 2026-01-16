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
        var deleteParams = ExtractDeleteParameters(parameters);

        var workbook = context.Document;
        var worksheet = ExcelHelper.GetWorksheet(workbook, deleteParams.SheetIndex);
        var chart = ExcelChartHelper.GetChart(worksheet, deleteParams.ChartIndex);

        var chartName = chart.Name ?? $"Chart {deleteParams.ChartIndex}";
        worksheet.Charts.RemoveAt(deleteParams.ChartIndex);

        MarkModified(context);

        return Success($"Chart #{deleteParams.ChartIndex} ({chartName}) deleted, {worksheet.Charts.Count} remaining");
    }

    /// <summary>
    ///     Extracts delete parameters from operation parameters.
    /// </summary>
    /// <param name="parameters">The operation parameters.</param>
    /// <returns>The extracted delete parameters.</returns>
    private static DeleteParameters ExtractDeleteParameters(OperationParameters parameters)
    {
        return new DeleteParameters(
            parameters.GetOptional("sheetIndex", 0),
            parameters.GetOptional("chartIndex", 0)
        );
    }

    /// <summary>
    ///     Record to hold delete chart parameters.
    /// </summary>
    private record DeleteParameters(int SheetIndex, int ChartIndex);
}
