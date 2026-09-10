using Aspose.Cells;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Helpers.Excel;
using AsposeMcpServer.Results.Excel.FreezePanes;

namespace AsposeMcpServer.Handlers.Excel.FreezePanes;

/// <summary>
///     Handler for getting freeze panes status from Excel worksheets.
/// </summary>
[ResultType(typeof(GetFreezePanesResult))]
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
    public override object Execute(OperationContext<Workbook> context, OperationParameters parameters)
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
            // The split cell is reported as-is. The previous -1 existed only to undo the +1 the
            // freeze handler applied when writing; with both operations on the same contract the
            // value read back is the value the caller asked for.
            worksheet.GetFreezedPanes(out var r, out var col, out var rows, out var cols);
            frozenRow = r;
            frozenColumn = col;
            frozenRows = rows;
            frozenColumns = cols;
        }

        return new GetFreezePanesResult
        {
            WorksheetName = worksheet.Name,
            IsFrozen = isFrozen,
            FrozenRow = frozenRow,
            FrozenColumn = frozenColumn,
            FrozenRows = frozenRows,
            FrozenColumns = frozenColumns,
            Status = isFrozen ? "Panes are frozen" : "Panes are not frozen"
        };
    }

    private static GetFreezePanesParameters ExtractGetFreezePanesParameters(OperationParameters parameters)
    {
        var sheetIndex = parameters.GetOptional("sheetIndex", 0);

        return new GetFreezePanesParameters(sheetIndex);
    }

    private sealed record GetFreezePanesParameters(int SheetIndex);
}
