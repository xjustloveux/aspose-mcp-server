using Aspose.Cells;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Helpers;

namespace AsposeMcpServer.Handlers.Excel.FreezePanes;

/// <summary>
///     Handler for getting freeze panes status from Excel worksheets.
/// </summary>
public class GetExcelFreezePanesHandler : OperationHandlerBase<Workbook>
{
    /// <inheritdoc />
    public override string Operation => "get";

    /// <summary>
    ///     Gets the current freeze panes status.
    /// </summary>
    /// <param name="context">The workbook context.</param>
    /// <param name="parameters">
    ///     Optional: sheetIndex (default: 0)
    /// </param>
    /// <returns>JSON result with freeze panes status information.</returns>
    public override string Execute(OperationContext<Workbook> context, OperationParameters parameters)
    {
        var p = ExtractGetFreezePanesParameters(parameters);

        var workbook = context.Document;
        var worksheet = ExcelHelper.GetWorksheet(workbook, p.SheetIndex);

        var isFrozen = worksheet.PaneState == PaneStateType.Frozen;
        int? frozenRow = null;
        int? frozenColumn = null;
        int? frozenRows = null;
        int? frozenColumns = null;

        if (isFrozen)
        {
            worksheet.GetFreezedPanes(out var r, out var col, out var rows, out var cols);
            frozenRow = r > 0 ? r - 1 : 0;
            frozenColumn = col > 0 ? col - 1 : 0;
            frozenRows = rows;
            frozenColumns = cols;
        }

        return JsonResult(new
        {
            worksheetName = worksheet.Name,
            isFrozen,
            frozenRow,
            frozenColumn,
            frozenRows,
            frozenColumns,
            status = isFrozen ? "Panes are frozen" : "Panes are not frozen"
        });
    }

    private static GetFreezePanesParameters ExtractGetFreezePanesParameters(OperationParameters parameters)
    {
        var sheetIndex = parameters.GetOptional("sheetIndex", 0);

        return new GetFreezePanesParameters(sheetIndex);
    }

    private record GetFreezePanesParameters(int SheetIndex);
}
