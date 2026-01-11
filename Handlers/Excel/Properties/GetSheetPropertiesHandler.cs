using Aspose.Cells;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Helpers;

namespace AsposeMcpServer.Handlers.Excel.Properties;

/// <summary>
///     Handler for getting worksheet properties from Excel files.
/// </summary>
public class GetSheetPropertiesHandler : OperationHandlerBase<Workbook>
{
    /// <inheritdoc />
    public override string Operation => "get_sheet_properties";

    /// <summary>
    ///     Gets worksheet properties including data counts, protection status, and print settings.
    /// </summary>
    /// <param name="context">The workbook context.</param>
    /// <param name="parameters">
    ///     Required: sheetIndex (0-based)
    /// </param>
    /// <returns>JSON result with worksheet properties.</returns>
    public override string Execute(OperationContext<Workbook> context, OperationParameters parameters)
    {
        var sheetIndex = parameters.GetOptional("sheetIndex", 0);

        var workbook = context.Document;
        var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);
        var pageSetup = worksheet.PageSetup;

        var result = new
        {
            name = worksheet.Name,
            index = sheetIndex,
            isVisible = worksheet.IsVisible,
            tabColor = worksheet.TabColor.ToString(),
            isSelected = workbook.Worksheets.ActiveSheetIndex == sheetIndex,
            dataRowCount = worksheet.Cells.MaxDataRow + 1,
            dataColumnCount = worksheet.Cells.MaxDataColumn + 1,
            isProtected = worksheet.Protection.IsProtectedWithPassword,
            commentsCount = worksheet.Comments.Count,
            chartsCount = worksheet.Charts.Count,
            picturesCount = worksheet.Pictures.Count,
            hyperlinksCount = worksheet.Hyperlinks.Count,
            printSettings = new
            {
                printArea = pageSetup.PrintArea,
                printTitleRows = pageSetup.PrintTitleRows,
                printTitleColumns = pageSetup.PrintTitleColumns,
                orientation = pageSetup.Orientation.ToString(),
                paperSize = pageSetup.PaperSize.ToString(),
                fitToPagesWide = pageSetup.FitToPagesWide,
                fitToPagesTall = pageSetup.FitToPagesTall
            }
        };

        return JsonResult(result);
    }
}
