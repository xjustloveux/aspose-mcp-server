using Aspose.Cells;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Helpers.Excel;
using AsposeMcpServer.Results.Excel.Chart;

namespace AsposeMcpServer.Handlers.Excel.Chart;

/// <summary>
///     Handler for getting charts from Excel worksheets.
/// </summary>
[ResultType(typeof(GetChartsResult))]
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
    public override object Execute(OperationContext<Workbook> context, OperationParameters parameters)
    {
        var getParams = ExtractGetParameters(parameters);

        var workbook = context.Document;
        var worksheet = ExcelHelper.GetWorksheet(workbook, getParams.SheetIndex);
        var charts = worksheet.Charts;

        if (charts.Count == 0)
            return new GetChartsResult
            {
                Count = 0,
                WorksheetName = worksheet.Name,
                Items = [],
                Message = "No charts found"
            };

        List<ChartInfo> chartList = [];
        for (var i = 0; i < charts.Count; i++)
        {
            var chart = charts[i];
            List<ChartSeriesInfo> seriesList = [];

            if (chart.NSeries is { Count: > 0 })
                for (var j = 0; j < chart.NSeries.Count && j < 5; j++)
                {
                    var series = chart.NSeries[j];
                    seriesList.Add(new ChartSeriesInfo
                    {
                        Index = j,
                        Name = series.Name ?? "(no name)",
                        ValuesRange = series.Values ?? "",
                        CategoryData = chart.NSeries.CategoryData
                    });
                }

            chartList.Add(new ChartInfo
            {
                Index = i,
                Name = chart.Name ?? "(no name)",
                Type = chart.Type.ToString(),
                Location = new ChartLocation
                {
                    UpperLeftRow = chart.ChartObject.UpperLeftRow,
                    LowerRightRow = chart.ChartObject.LowerRightRow,
                    UpperLeftColumn = chart.ChartObject.UpperLeftColumn,
                    LowerRightColumn = chart.ChartObject.LowerRightColumn
                },
                Width = chart.ChartObject.Width,
                Height = chart.ChartObject.Height,
                Title = chart.Title?.Text,
                LegendEnabled = chart.Legend != null,
                SeriesCount = chart.NSeries?.Count ?? 0,
                Series = seriesList
            });
        }

        return new GetChartsResult
        {
            Count = charts.Count,
            WorksheetName = worksheet.Name,
            Items = chartList
        };
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
    private sealed record GetParameters(int SheetIndex);
}
