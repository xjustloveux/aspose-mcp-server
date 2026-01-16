using Aspose.Cells;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Helpers;

namespace AsposeMcpServer.Handlers.Excel.Protect;

/// <summary>
///     Handler for setting cell locked status in Excel worksheet.
/// </summary>
public class SetExcelCellLockedHandler : OperationHandlerBase<Workbook>
{
    /// <inheritdoc />
    public override string Operation => "set_cell_locked";

    /// <summary>
    ///     Sets cell locked status.
    /// </summary>
    /// <param name="context">The workbook context.</param>
    /// <param name="parameters">
    ///     Required: range
    ///     Optional: sheetIndex (default: 0), locked (default: false)
    /// </param>
    /// <returns>Success message with cell lock status details.</returns>
    /// <exception cref="ArgumentException">Thrown when range is empty or null.</exception>
    public override string Execute(OperationContext<Workbook> context, OperationParameters parameters)
    {
        var p = ExtractSetExcelCellLockedParameters(parameters);

        if (string.IsNullOrEmpty(p.Range))
            throw new ArgumentException("range is required for set_cell_locked operation");

        var workbook = context.Document;
        var worksheet = ExcelHelper.GetWorksheet(workbook, p.SheetIndex);
        var cells = worksheet.Cells;

        var cellRange = ExcelHelper.CreateRange(cells, p.Range);

        var style = workbook.CreateStyle();
        style.IsLocked = p.Locked;

        var styleFlag = new StyleFlag { Locked = true };
        cellRange.ApplyStyle(style, styleFlag);

        MarkModified(context);

        return Success(
            $"Cell lock status set to {(p.Locked ? "locked" : "unlocked")} for range {p.Range} in sheet {p.SheetIndex}.");
    }

    /// <summary>
    ///     Extracts parameters for SetExcelCellLocked operation.
    /// </summary>
    /// <param name="parameters">The operation parameters.</param>
    /// <returns>Extracted parameters.</returns>
    private static SetExcelCellLockedParameters ExtractSetExcelCellLockedParameters(OperationParameters parameters)
    {
        return new SetExcelCellLockedParameters(
            parameters.GetRequired<string>("range"),
            parameters.GetOptional("sheetIndex", 0),
            parameters.GetOptional("locked", false)
        );
    }

    /// <summary>
    ///     Parameters for SetExcelCellLocked operation.
    /// </summary>
    /// <param name="Range">The cell range to lock/unlock.</param>
    /// <param name="SheetIndex">The sheet index.</param>
    /// <param name="Locked">Whether to lock the cells.</param>
    private sealed record SetExcelCellLockedParameters(string Range, int SheetIndex, bool Locked);
}
