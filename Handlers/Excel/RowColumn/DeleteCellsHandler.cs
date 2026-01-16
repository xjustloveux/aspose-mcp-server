using Aspose.Cells;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Helpers;

namespace AsposeMcpServer.Handlers.Excel.RowColumn;

/// <summary>
///     Handler for deleting cells in a range and shifting remaining cells.
/// </summary>
public class DeleteCellsHandler : OperationHandlerBase<Workbook>
{
    /// <inheritdoc />
    public override string Operation => "delete_cells";

    /// <summary>
    ///     Deletes cells in a range and shifts remaining cells.
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="parameters">
    ///     Required: range, shiftDirection
    ///     Optional: sheetIndex (default: 0)
    /// </param>
    /// <returns>Success message with deletion details.</returns>
    public override string Execute(OperationContext<Workbook> context, OperationParameters parameters)
    {
        var p = ExtractDeleteCellsParameters(parameters);

        var workbook = context.Document;
        var worksheet = ExcelHelper.GetWorksheet(workbook, p.SheetIndex);

        var rangeObj = ExcelHelper.CreateRange(worksheet.Cells, p.Range);
        var shiftType = string.Equals(p.ShiftDirection, "left", StringComparison.OrdinalIgnoreCase)
            ? ShiftType.Left
            : ShiftType.Up;

        worksheet.Cells.DeleteRange(
            rangeObj.FirstRow,
            rangeObj.FirstColumn,
            rangeObj.FirstRow + rangeObj.RowCount - 1,
            rangeObj.FirstColumn + rangeObj.ColumnCount - 1,
            shiftType);

        MarkModified(context);

        return Success($"Cells deleted in range {p.Range}, shifted {p.ShiftDirection}.");
    }

    private static DeleteCellsParameters ExtractDeleteCellsParameters(OperationParameters parameters)
    {
        var sheetIndex = parameters.GetOptional("sheetIndex", 0);
        var range = parameters.GetRequired<string>("range");
        var shiftDirection = parameters.GetRequired<string>("shiftDirection");

        return new DeleteCellsParameters(sheetIndex, range, shiftDirection);
    }

    private record DeleteCellsParameters(int SheetIndex, string Range, string ShiftDirection);
}
