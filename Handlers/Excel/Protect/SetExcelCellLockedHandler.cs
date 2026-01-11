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
    public override string Execute(OperationContext<Workbook> context, OperationParameters parameters)
    {
        var range = parameters.GetRequired<string>("range");
        var sheetIndex = parameters.GetOptional("sheetIndex", 0);
        var locked = parameters.GetOptional("locked", false);

        if (string.IsNullOrEmpty(range))
            throw new ArgumentException("range is required for set_cell_locked operation");

        var workbook = context.Document;
        var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);
        var cells = worksheet.Cells;

        var cellRange = ExcelHelper.CreateRange(cells, range);

        var style = workbook.CreateStyle();
        style.IsLocked = locked;

        var styleFlag = new StyleFlag { Locked = true };
        cellRange.ApplyStyle(style, styleFlag);

        MarkModified(context);

        return Success(
            $"Cell lock status set to {(locked ? "locked" : "unlocked")} for range {range} in sheet {sheetIndex}.");
    }
}
