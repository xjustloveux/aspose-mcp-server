using Aspose.Cells;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Helpers;

namespace AsposeMcpServer.Handlers.Excel.Protect;

/// <summary>
///     Handler for getting protection status of Excel workbook or worksheet.
/// </summary>
public class GetExcelProtectionHandler : OperationHandlerBase<Workbook>
{
    /// <inheritdoc />
    public override string Operation => "get";

    /// <summary>
    ///     Gets protection status for workbook or worksheet.
    /// </summary>
    /// <param name="context">The workbook context.</param>
    /// <param name="parameters">
    ///     Optional: sheetIndex (if not specified, returns all worksheets' protection status)
    /// </param>
    /// <returns>JSON string containing protection status information.</returns>
    public override string Execute(OperationContext<Workbook> context, OperationParameters parameters)
    {
        var sheetIndex = parameters.GetOptional<int?>("sheetIndex");

        var workbook = context.Document;
        List<object> worksheets = [];

        if (sheetIndex.HasValue)
        {
            var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex.Value);
            worksheets.Add(CreateSheetProtectionInfo(worksheet, sheetIndex.Value));
        }
        else
        {
            for (var i = 0; i < workbook.Worksheets.Count; i++)
                worksheets.Add(CreateSheetProtectionInfo(workbook.Worksheets[i], i));
        }

        var result = new
        {
            count = worksheets.Count,
            totalWorksheets = workbook.Worksheets.Count,
            worksheets
        };

        return JsonResult(result);
    }

    /// <summary>
    ///     Creates protection information object for a worksheet.
    /// </summary>
    /// <param name="worksheet">The worksheet to get protection information from.</param>
    /// <param name="index">The index of the worksheet.</param>
    /// <returns>An anonymous object containing the worksheet's protection details.</returns>
    private static object CreateSheetProtectionInfo(Worksheet worksheet, int index)
    {
        var protection = worksheet.Protection;
        return new
        {
            index,
            name = worksheet.Name,
            isProtected = protection.IsProtectedWithPassword,
            allowSelectingLockedCell = protection.AllowSelectingLockedCell,
            allowSelectingUnlockedCell = protection.AllowSelectingUnlockedCell,
            allowFormattingCell = protection.AllowFormattingCell,
            allowFormattingColumn = protection.AllowFormattingColumn,
            allowFormattingRow = protection.AllowFormattingRow,
            allowInsertingColumn = protection.AllowInsertingColumn,
            allowInsertingRow = protection.AllowInsertingRow,
            allowInsertingHyperlink = protection.AllowInsertingHyperlink,
            allowDeletingColumn = protection.AllowDeletingColumn,
            allowDeletingRow = protection.AllowDeletingRow,
            allowSorting = protection.AllowSorting,
            allowFiltering = protection.AllowFiltering,
            allowUsingPivotTable = protection.AllowUsingPivotTable,
            allowEditingObject = protection.AllowEditingObject,
            allowEditingScenario = protection.AllowEditingScenario
        };
    }
}
