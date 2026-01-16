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
    /// <exception cref="ArgumentException">Thrown when dataRange is not provided.</exception>
    public override string Execute(OperationContext<Workbook> context, OperationParameters parameters)
    {
        var addParams = ExtractAddParameters(parameters);

        if (string.IsNullOrEmpty(addParams.DataRange))
            throw new ArgumentException("dataRange is required for add operation");

        var workbook = context.Document;
        var worksheet = ExcelHelper.GetWorksheet(workbook, addParams.SheetIndex);
        var chartType = ExcelChartHelper.ParseChartType(addParams.ChartType);

        int chartTopRow;
        if (addParams.TopRow.HasValue)
        {
            chartTopRow = addParams.TopRow.Value;
        }
        else
        {
            var dataRangeObj = ExcelHelper.CreateRange(worksheet.Cells, addParams.DataRange.Split(',')[0].Trim());
            chartTopRow = dataRangeObj.FirstRow + dataRangeObj.RowCount + 2;
        }

        var chartIndex =
            worksheet.Charts.Add(chartType, chartTopRow, addParams.LeftColumn, chartTopRow + addParams.Height,
                addParams.LeftColumn + addParams.Width);
        var chart = worksheet.Charts[chartIndex];

        ExcelChartHelper.AddDataSeries(chart, addParams.DataRange);
        ExcelChartHelper.SetCategoryData(chart, addParams.CategoryAxisDataRange ?? "");

        if (!string.IsNullOrEmpty(addParams.Title))
            chart.Title.Text = addParams.Title;

        workbook.CalculateFormula();

        MarkModified(context);

        var result = $"Chart added with data range: {addParams.DataRange}";
        if (!string.IsNullOrEmpty(addParams.CategoryAxisDataRange))
            result += $", X-axis: {addParams.CategoryAxisDataRange}";
        return Success(result);
    }

    /// <summary>
    ///     Extracts add parameters from operation parameters.
    /// </summary>
    /// <param name="parameters">The operation parameters.</param>
    /// <returns>The extracted add parameters.</returns>
    private static AddParameters ExtractAddParameters(OperationParameters parameters)
    {
        return new AddParameters(
            parameters.GetOptional("sheetIndex", 0),
            parameters.GetOptional<string?>("chartType"),
            parameters.GetOptional<string?>("dataRange"),
            parameters.GetOptional<string?>("categoryAxisDataRange"),
            parameters.GetOptional<string?>("title"),
            parameters.GetOptional<int?>("topRow"),
            parameters.GetOptional("leftColumn", 0),
            parameters.GetOptional("width", 10),
            parameters.GetOptional("height", 15)
        );
    }

    /// <summary>
    ///     Record to hold add chart parameters.
    /// </summary>
    private sealed record AddParameters(
        int SheetIndex,
        string? ChartType,
        string? DataRange,
        string? CategoryAxisDataRange,
        string? Title,
        int? TopRow,
        int LeftColumn,
        int Width,
        int Height);
}
