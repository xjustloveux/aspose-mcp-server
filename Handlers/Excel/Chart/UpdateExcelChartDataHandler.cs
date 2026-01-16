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
    /// <exception cref="ArgumentException">Thrown when dataRange is not provided.</exception>
    public override string Execute(OperationContext<Workbook> context, OperationParameters parameters)
    {
        var updateParams = ExtractUpdateDataParameters(parameters);

        if (string.IsNullOrEmpty(updateParams.DataRange))
            throw new ArgumentException("dataRange is required for update_data operation");

        var workbook = context.Document;
        var worksheet = ExcelHelper.GetWorksheet(workbook, updateParams.SheetIndex);
        var chart = ExcelChartHelper.GetChart(worksheet, updateParams.ChartIndex);

        ExcelChartHelper.AddDataSeries(chart, updateParams.DataRange);
        ExcelChartHelper.SetCategoryData(chart, updateParams.CategoryAxisDataRange ?? "");

        MarkModified(context);

        var result = $"Chart #{updateParams.ChartIndex} data updated to: {updateParams.DataRange}";
        if (!string.IsNullOrEmpty(updateParams.CategoryAxisDataRange))
            result += $", X-axis: {updateParams.CategoryAxisDataRange}";
        return Success(result);
    }

    /// <summary>
    ///     Extracts update data parameters from operation parameters.
    /// </summary>
    /// <param name="parameters">The operation parameters.</param>
    /// <returns>The extracted update data parameters.</returns>
    private static UpdateDataParameters ExtractUpdateDataParameters(OperationParameters parameters)
    {
        return new UpdateDataParameters(
            parameters.GetOptional("sheetIndex", 0),
            parameters.GetOptional("chartIndex", 0),
            parameters.GetOptional<string?>("dataRange"),
            parameters.GetOptional<string?>("categoryAxisDataRange")
        );
    }

    /// <summary>
    ///     Record to hold update data parameters.
    /// </summary>
    private record UpdateDataParameters(
        int SheetIndex,
        int ChartIndex,
        string? DataRange,
        string? CategoryAxisDataRange);
}
