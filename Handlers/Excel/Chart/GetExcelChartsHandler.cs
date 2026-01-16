using Aspose.Cells;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Helpers;

namespace AsposeMcpServer.Handlers.Excel.Chart;

/// <summary>
///     Handler for getting charts from Excel worksheets.
/// </summary>
public class GetExcelChartsHandler : OperationHandlerBase<Workbook>
{
    /// <inheritdoc />
    public override string Operation => "get";

    /// <summary>
    ///     Gets all charts from the worksheet.
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="parameters">
    ///     Optional: sheetIndex
    /// </param>
    /// <returns>JSON string containing chart information.</returns>
    public override string Execute(OperationContext<Workbook> context, OperationParameters parameters)
    {
        var getParams = ExtractGetParameters(parameters);

        var workbook = context.Document;
        var worksheet = ExcelHelper.GetWorksheet(workbook, getParams.SheetIndex);
        var charts = worksheet.Charts;

        if (charts.Count == 0)
        {
            var emptyResult = new
            {
                count = 0,
                worksheetName = worksheet.Name,
                items = Array.Empty<object>(),
                message = "No charts found"
            };
            return JsonResult(emptyResult);
        }

        List<object> chartList = [];
        for (var i = 0; i < charts.Count; i++)
        {
            var chart = charts[i];
            List<object> seriesList = [];

            if (chart.NSeries is { Count: > 0 })
                for (var j = 0; j < chart.NSeries.Count && j < 5; j++)
                {
                    var series = chart.NSeries[j];
                    seriesList.Add(new
                    {
                        index = j,
                        name = series.Name ?? "(no name)",
                        valuesRange = series.Values ?? "",
                        categoryData = chart.NSeries.CategoryData
                    });
                }

            chartList.Add(new
            {
                index = i,
                name = chart.Name ?? "(no name)",
                type = chart.Type.ToString(),
                location = new
                {
                    upperLeftRow = chart.ChartObject.UpperLeftRow,
                    lowerRightRow = chart.ChartObject.LowerRightRow,
                    upperLeftColumn = chart.ChartObject.UpperLeftColumn,
                    lowerRightColumn = chart.ChartObject.LowerRightColumn
                },
                width = chart.ChartObject.Width,
                height = chart.ChartObject.Height,
                title = chart.Title?.Text,
                legendEnabled = chart.Legend != null,
                seriesCount = chart.NSeries?.Count ?? 0,
                series = seriesList
            });
        }

        var result = new
        {
            count = charts.Count,
            worksheetName = worksheet.Name,
            items = chartList
        };

        return JsonResult(result);
    }

    /// <summary>
    ///     Extracts get parameters from operation parameters.
    /// </summary>
    /// <param name="parameters">The operation parameters.</param>
    /// <returns>The extracted get parameters.</returns>
    private static GetParameters ExtractGetParameters(OperationParameters parameters)
    {
        return new GetParameters(parameters.GetOptional("sheetIndex", 0));
    }

    /// <summary>
    ///     Record to hold get charts parameters.
    /// </summary>
    private record GetParameters(int SheetIndex);
}
