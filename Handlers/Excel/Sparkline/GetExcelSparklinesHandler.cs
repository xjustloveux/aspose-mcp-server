using Aspose.Cells;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Helpers.Excel;
using AsposeMcpServer.Results.Excel.Sparkline;

namespace AsposeMcpServer.Handlers.Excel.Sparkline;

/// <summary>
///     Handler for getting sparkline groups from an Excel worksheet.
/// </summary>
[ResultType(typeof(GetSparklinesExcelResult))]
public class GetExcelSparklinesHandler : OperationHandlerBase<Workbook>
{
    /// <inheritdoc />
    public override string Operation => "get";

    /// <summary>
    ///     Gets all sparkline groups from a worksheet.
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="parameters">
    ///     Optional: sheetIndex (default: 0)
    /// </param>
    /// <returns>Sparkline group information result.</returns>
    /// <exception cref="ArgumentException">Thrown when parameters are invalid.</exception>
    public override object Execute(OperationContext<Workbook> context, OperationParameters parameters)
    {
        var sheetIndex = parameters.GetOptional("sheetIndex", 0);

        try
        {
            var workbook = context.Document;
            var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);

            var items = new List<ExcelSparklineInfo>();
            for (var i = 0; i < worksheet.SparklineGroups.Count; i++)
            {
                var group = worksheet.SparklineGroups[i];
                string? dataRange = null;
                string? locationRange = null;

                if (group.Sparklines.Count > 0)
                {
                    dataRange = group.Sparklines[0].DataRange;
                    locationRange =
                        $"R{group.Sparklines[0].Row + 1}C{group.Sparklines[0].Column + 1}";
                }

                items.Add(new ExcelSparklineInfo
                {
                    Index = i,
                    Type = group.Type.ToString(),
                    SparklineCount = group.Sparklines.Count,
                    DataRange = dataRange,
                    LocationRange = locationRange
                });
            }

            return new GetSparklinesExcelResult
            {
                Count = items.Count,
                SheetIndex = sheetIndex,
                Items = items,
                Message = items.Count == 0 ? "No sparkline groups found in the worksheet." : null
            };
        }
        catch (CellsException ex)
        {
            throw new ArgumentException($"Failed to get sparklines from sheet {sheetIndex}: {ex.Message}");
        }
    }
}
