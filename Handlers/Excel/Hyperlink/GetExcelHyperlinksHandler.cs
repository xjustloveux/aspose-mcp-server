using Aspose.Cells;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Helpers.Excel;
using AsposeMcpServer.Results.Excel.Hyperlink;

namespace AsposeMcpServer.Handlers.Excel.Hyperlink;

/// <summary>
///     Handler for getting hyperlinks from Excel worksheets.
/// </summary>
[ResultType(typeof(GetHyperlinksExcelResult))]
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
    public override object Execute(OperationContext<Workbook> context, OperationParameters parameters)
    {
        var getParams = ExtractGetParameters(parameters);

        var workbook = context.Document;
        var worksheet = ExcelHelper.GetWorksheet(workbook, getParams.SheetIndex);
        var hyperlinks = worksheet.Hyperlinks;

        if (hyperlinks.Count == 0)
            return new GetHyperlinksExcelResult
            {
                Count = 0,
                WorksheetName = worksheet.Name,
                Items = Array.Empty<ExcelHyperlinkInfo>(),
                Message = "No hyperlinks found"
            };

        List<ExcelHyperlinkInfo> hyperlinkList = [];
        for (var i = 0; i < hyperlinks.Count; i++)
        {
            var hyperlink = hyperlinks[i];
            var area = hyperlink.Area;
            var cellRef = CellsHelper.CellIndexToName(area.StartRow, area.StartColumn);

            hyperlinkList.Add(new ExcelHyperlinkInfo
            {
                Index = i,
                Cell = cellRef,
                Url = hyperlink.Address,
                DisplayText = hyperlink.TextToDisplay,
                Area = new HyperlinkArea
                {
                    StartCell = cellRef,
                    EndCell = CellsHelper.CellIndexToName(area.EndRow, area.EndColumn)
                }
            });
        }

        return new GetHyperlinksExcelResult
        {
            Count = hyperlinks.Count,
            WorksheetName = worksheet.Name,
            Items = hyperlinkList
        };
    }

    private static GetParameters ExtractGetParameters(OperationParameters parameters)
    {
        var sheetIndex = parameters.GetOptional("sheetIndex", 0);

        return new GetParameters(sheetIndex);
    }

    private sealed record GetParameters(int SheetIndex);
}
