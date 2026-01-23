using Aspose.Cells;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Helpers.Excel;
using AsposeMcpServer.Results.Excel.Protect;

namespace AsposeMcpServer.Handlers.Excel.Protect;

/// <summary>
///     Handler for getting protection status of Excel workbook or worksheet.
/// </summary>
[ResultType(typeof(GetProtectionResult))]
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
    public override object Execute(OperationContext<Workbook> context, OperationParameters parameters)
    {
        var p = ExtractGetExcelProtectionParameters(parameters);

        var workbook = context.Document;
        List<WorksheetProtectionInfo> worksheets = [];

        if (p.SheetIndex.HasValue)
        {
            var worksheet = ExcelHelper.GetWorksheet(workbook, p.SheetIndex.Value);
            worksheets.Add(CreateSheetProtectionInfo(worksheet, p.SheetIndex.Value));
        }
        else
        {
            for (var i = 0; i < workbook.Worksheets.Count; i++)
                worksheets.Add(CreateSheetProtectionInfo(workbook.Worksheets[i], i));
        }

        return new GetProtectionResult
        {
            Count = worksheets.Count,
            TotalWorksheets = workbook.Worksheets.Count,
            Worksheets = worksheets
        };
    }

    /// <summary>
    ///     Creates protection information object for a worksheet.
    /// </summary>
    /// <param name="worksheet">The worksheet to get protection information from.</param>
    /// <param name="index">The index of the worksheet.</param>
    /// <returns>WorksheetProtectionInfo containing the worksheet's protection details.</returns>
    private static WorksheetProtectionInfo CreateSheetProtectionInfo(Worksheet worksheet, int index)
    {
        var protection = worksheet.Protection;
        return new WorksheetProtectionInfo
        {
            Index = index,
            Name = worksheet.Name,
            IsProtected = protection.IsProtectedWithPassword,
            AllowSelectingLockedCell = protection.AllowSelectingLockedCell,
            AllowSelectingUnlockedCell = protection.AllowSelectingUnlockedCell,
            AllowFormattingCell = protection.AllowFormattingCell,
            AllowFormattingColumn = protection.AllowFormattingColumn,
            AllowFormattingRow = protection.AllowFormattingRow,
            AllowInsertingColumn = protection.AllowInsertingColumn,
            AllowInsertingRow = protection.AllowInsertingRow,
            AllowInsertingHyperlink = protection.AllowInsertingHyperlink,
            AllowDeletingColumn = protection.AllowDeletingColumn,
            AllowDeletingRow = protection.AllowDeletingRow,
            AllowSorting = protection.AllowSorting,
            AllowFiltering = protection.AllowFiltering,
            AllowUsingPivotTable = protection.AllowUsingPivotTable,
            AllowEditingObject = protection.AllowEditingObject,
            AllowEditingScenario = protection.AllowEditingScenario
        };
    }

    /// <summary>
    ///     Extracts parameters for GetExcelProtection operation.
    /// </summary>
    /// <param name="parameters">The operation parameters.</param>
    /// <returns>Extracted parameters.</returns>
    private static GetExcelProtectionParameters ExtractGetExcelProtectionParameters(OperationParameters parameters)
    {
        return new GetExcelProtectionParameters(
            parameters.GetOptional<int?>("sheetIndex")
        );
    }

    /// <summary>
    ///     Parameters for GetExcelProtection operation.
    /// </summary>
    /// <param name="SheetIndex">The sheet index (optional).</param>
    private sealed record GetExcelProtectionParameters(int? SheetIndex);
}
