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
        var getParams = ExtractGetSheetPropertiesParameters(parameters);

        var workbook = context.Document;
        var worksheet = ExcelHelper.GetWorksheet(workbook, getParams.SheetIndex);
        var pageSetup = worksheet.PageSetup;

        var result = new
        {
            name = worksheet.Name,
            index = getParams.SheetIndex,
            isVisible = worksheet.IsVisible,
            tabColor = worksheet.TabColor.ToString(),
            isSelected = workbook.Worksheets.ActiveSheetIndex == getParams.SheetIndex,
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

    /// <summary>
    ///     Extracts get sheet properties parameters from operation parameters.
    /// </summary>
    /// <param name="parameters">The operation parameters.</param>
    /// <returns>The extracted get sheet properties parameters.</returns>
    private static GetSheetPropertiesParameters ExtractGetSheetPropertiesParameters(OperationParameters parameters)
    {
        return new GetSheetPropertiesParameters(
            parameters.GetOptional("sheetIndex", 0)
        );
    }

    /// <summary>
    ///     Record to hold get sheet properties parameters.
    /// </summary>
    private record GetSheetPropertiesParameters(int SheetIndex);
}
