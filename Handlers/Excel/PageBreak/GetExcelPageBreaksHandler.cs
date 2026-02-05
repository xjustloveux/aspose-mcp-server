using Aspose.Cells;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Helpers.Excel;
using AsposeMcpServer.Results.Excel.PageBreak;

namespace AsposeMcpServer.Handlers.Excel.PageBreak;

/// <summary>
///     Handler for getting page breaks from an Excel worksheet.
/// </summary>
[ResultType(typeof(GetPageBreaksExcelResult))]
public class GetExcelPageBreaksHandler : OperationHandlerBase<Workbook>
{
    /// <inheritdoc />
    public override string Operation => "get";

    /// <summary>
    ///     Gets all page breaks from a worksheet.
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="parameters">
    ///     Optional: sheetIndex (default: 0)
    /// </param>
    /// <returns>Page break information result.</returns>
    /// <exception cref="ArgumentException">Thrown when parameters are invalid.</exception>
    public override object Execute(OperationContext<Workbook> context, OperationParameters parameters)
    {
        var sheetIndex = parameters.GetOptional("sheetIndex", 0);

        try
        {
            var workbook = context.Document;
            var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);

            var items = new List<ExcelPageBreakInfo>();

            for (var i = 0; i < worksheet.HorizontalPageBreaks.Count; i++)
            {
                var pb = worksheet.HorizontalPageBreaks[i];
                items.Add(new ExcelPageBreakInfo
                {
                    Index = i,
                    Type = "Horizontal",
                    StartIndex = pb.Row,
                    EndIndex = pb.Row
                });
            }

            for (var i = 0; i < worksheet.VerticalPageBreaks.Count; i++)
            {
                var pb = worksheet.VerticalPageBreaks[i];
                items.Add(new ExcelPageBreakInfo
                {
                    Index = i,
                    Type = "Vertical",
                    StartIndex = pb.Column,
                    EndIndex = pb.Column
                });
            }

            return new GetPageBreaksExcelResult
            {
                Count = items.Count,
                SheetIndex = sheetIndex,
                Items = items,
                Message = items.Count == 0 ? "No page breaks found in the worksheet." : null
            };
        }
        catch (CellsException ex)
        {
            throw new ArgumentException($"Failed to get page breaks from sheet {sheetIndex}: {ex.Message}");
        }
    }
}
