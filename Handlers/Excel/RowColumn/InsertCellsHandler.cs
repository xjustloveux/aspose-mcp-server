using Aspose.Cells;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Helpers.Excel;
using AsposeMcpServer.Results.Common;

namespace AsposeMcpServer.Handlers.Excel.RowColumn;

/// <summary>
///     Handler for inserting cells in a range and shifting existing cells.
/// </summary>
[ResultType(typeof(SuccessResult))]
public class InsertCellsHandler : OperationHandlerBase<Workbook>
{
    /// <inheritdoc />
    public override string Operation => "insert_cells";

    /// <summary>
    ///     Inserts cells in a range and shifts existing cells.
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="parameters">
    ///     Required: range, shiftDirection
    ///     Optional: sheetIndex (default: 0)
    /// </param>
    /// <returns>Success message with insertion details.</returns>
    public override object Execute(OperationContext<Workbook> context, OperationParameters parameters)
    {
        var p = ExtractInsertCellsParameters(parameters);

        var workbook = context.Document;
        var worksheet = ExcelHelper.GetWorksheet(workbook, p.SheetIndex);

        var rangeObj = ExcelHelper.CreateRange(worksheet.Cells, p.Range);
        var shiftType = string.Equals(p.ShiftDirection, "right", StringComparison.OrdinalIgnoreCase)
            ? ShiftType.Right
            : ShiftType.Down;

        var cellArea = CellArea.CreateCellArea(
            rangeObj.FirstRow,
            rangeObj.FirstColumn,
            rangeObj.FirstRow + rangeObj.RowCount - 1,
            rangeObj.FirstColumn + rangeObj.ColumnCount - 1);

        worksheet.Cells.InsertRange(cellArea, shiftType);

        MarkModified(context);

        return new SuccessResult { Message = $"Cells inserted in range {p.Range}, shifted {p.ShiftDirection}." };
    }

    private static InsertCellsParameters ExtractInsertCellsParameters(OperationParameters parameters)
    {
        var sheetIndex = parameters.GetOptional("sheetIndex", 0);
        var range = parameters.GetRequired<string>("range");
        var shiftDirection = parameters.GetRequired<string>("shiftDirection");

        return new InsertCellsParameters(sheetIndex, range, shiftDirection);
    }

    private sealed record InsertCellsParameters(int SheetIndex, string Range, string ShiftDirection);
}
