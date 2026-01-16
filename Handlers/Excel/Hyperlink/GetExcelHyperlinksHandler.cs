using Aspose.Cells;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Helpers;

namespace AsposeMcpServer.Handlers.Excel.Hyperlink;

/// <summary>
///     Handler for getting hyperlinks from Excel worksheets.
/// </summary>
public class GetExcelHyperlinksHandler : OperationHandlerBase<Workbook>
{
    /// <inheritdoc />
    public override string Operation => "get";

    /// <summary>
    ///     Gets all hyperlinks from the worksheet.
    /// </summary>
    /// <param name="context">The workbook context.</param>
    /// <param name="parameters">
    ///     Optional: sheetIndex (default: 0)
    /// </param>
    /// <returns>JSON result with hyperlink information.</returns>
    public override string Execute(OperationContext<Workbook> context, OperationParameters parameters)
    {
        var getParams = ExtractGetParameters(parameters);

        var workbook = context.Document;
        var worksheet = ExcelHelper.GetWorksheet(workbook, getParams.SheetIndex);
        var hyperlinks = worksheet.Hyperlinks;

        if (hyperlinks.Count == 0)
            return JsonResult(new
            {
                count = 0,
                worksheetName = worksheet.Name,
                items = Array.Empty<object>(),
                message = "No hyperlinks found"
            });

        List<object> hyperlinkList = [];
        for (var i = 0; i < hyperlinks.Count; i++)
        {
            var hyperlink = hyperlinks[i];
            var area = hyperlink.Area;
            var cellRef = CellsHelper.CellIndexToName(area.StartRow, area.StartColumn);

            hyperlinkList.Add(new
            {
                index = i,
                cell = cellRef,
                url = hyperlink.Address,
                displayText = hyperlink.TextToDisplay,
                area = new
                {
                    startCell = cellRef,
                    endCell = CellsHelper.CellIndexToName(area.EndRow, area.EndColumn)
                }
            });
        }

        return JsonResult(new
        {
            count = hyperlinks.Count,
            worksheetName = worksheet.Name,
            items = hyperlinkList
        });
    }

    private static GetParameters ExtractGetParameters(OperationParameters parameters)
    {
        var sheetIndex = parameters.GetOptional("sheetIndex", 0);

        return new GetParameters(sheetIndex);
    }

    private sealed record GetParameters(int SheetIndex);
}
