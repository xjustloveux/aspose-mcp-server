using Aspose.Cells;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Helpers;

namespace AsposeMcpServer.Handlers.Excel.Chart;

/// <summary>
///     Handler for adding charts to Excel worksheets.
/// </summary>
public class AddExcelChartHandler : OperationHandlerBase<Workbook>
{
    /// <inheritdoc />
    public override string Operation => "add";

    /// <summary>
    ///     Adds a new chart to the worksheet.
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="parameters">
    ///     Required: dataRange
    ///     Optional: sheetIndex, chartType, categoryAxisDataRange, title, topRow, leftColumn, width, height
    /// </param>
    /// <returns>Success message with chart creation details.</returns>
    public override string Execute(OperationContext<Workbook> context, OperationParameters parameters)
    {
        var sheetIndex = parameters.GetOptional("sheetIndex", 0);
        var chartTypeStr = parameters.GetOptional<string?>("chartType");
        var dataRange = parameters.GetOptional<string?>("dataRange");
        var categoryAxisDataRange = parameters.GetOptional<string?>("categoryAxisDataRange");
        var title = parameters.GetOptional<string?>("title");
        var topRow = parameters.GetOptional<int?>("topRow");
        var leftColumn = parameters.GetOptional("leftColumn", 0);
        var width = parameters.GetOptional("width", 10);
        var height = parameters.GetOptional("height", 15);

        if (string.IsNullOrEmpty(dataRange))
            throw new ArgumentException("dataRange is required for add operation");

        var workbook = context.Document;
        var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);
        var chartType = ExcelChartHelper.ParseChartType(chartTypeStr);

        int chartTopRow;
        if (topRow.HasValue)
        {
            chartTopRow = topRow.Value;
        }
        else
        {
            var dataRangeObj = ExcelHelper.CreateRange(worksheet.Cells, dataRange.Split(',')[0].Trim());
            chartTopRow = dataRangeObj.FirstRow + dataRangeObj.RowCount + 2;
        }

        var chartIndex =
            worksheet.Charts.Add(chartType, chartTopRow, leftColumn, chartTopRow + height, leftColumn + width);
        var chart = worksheet.Charts[chartIndex];

        ExcelChartHelper.AddDataSeries(chart, dataRange);
        ExcelChartHelper.SetCategoryData(chart, categoryAxisDataRange ?? "");

        if (!string.IsNullOrEmpty(title))
            chart.Title.Text = title;

        workbook.CalculateFormula();

        MarkModified(context);

        var result = $"Chart added with data range: {dataRange}";
        if (!string.IsNullOrEmpty(categoryAxisDataRange))
            result += $", X-axis: {categoryAxisDataRange}";
        return Success(result);
    }
}
