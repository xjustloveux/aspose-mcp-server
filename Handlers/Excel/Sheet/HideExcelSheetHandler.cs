using Aspose.Cells;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Helpers.Excel;
using AsposeMcpServer.Results.Common;

namespace AsposeMcpServer.Handlers.Excel.Sheet;

/// <summary>
///     Handler for hiding/showing worksheets in Excel workbooks.
/// </summary>
[ResultType(typeof(SuccessResult))]
public class HideExcelSheetHandler : OperationHandlerBase<Workbook>
{
    /// <inheritdoc />
    public override string Operation => "hide";

    /// <summary>
    ///     Toggles the visibility of a worksheet in the workbook.
    /// </summary>
    /// <param name="context">The workbook context.</param>
    /// <param name="parameters">
    ///     Required: sheetIndex (0-based index of sheet to hide/show)
    /// </param>
    /// <returns>Success message indicating whether the sheet was hidden or shown.</returns>
    public override object Execute(OperationContext<Workbook> context, OperationParameters parameters)
    {
        var p = ExtractHideExcelSheetParameters(parameters);

        var workbook = context.Document;
        var worksheet = ExcelHelper.GetWorksheet(workbook, p.SheetIndex);
        var sheetName = worksheet.Name;

        if (worksheet.IsVisible)
        {
            worksheet.IsVisible = false;
            MarkModified(context);
            return new SuccessResult { Message = $"Worksheet '{sheetName}' hidden." };
        }

        worksheet.IsVisible = true;
        MarkModified(context);
        return new SuccessResult { Message = $"Worksheet '{sheetName}' shown." };
    }

    private static HideExcelSheetParameters ExtractHideExcelSheetParameters(OperationParameters parameters)
    {
        var sheetIndex = parameters.GetRequired<int>("sheetIndex");

        return new HideExcelSheetParameters(sheetIndex);
    }

    private sealed record HideExcelSheetParameters(int SheetIndex);
}
