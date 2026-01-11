using Aspose.Cells;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Helpers;

namespace AsposeMcpServer.Handlers.Excel.Cell;

/// <summary>
///     Handler for editing Excel cells with value, formula, or clearing.
/// </summary>
public class EditExcelCellHandler : OperationHandlerBase<Workbook>
{
    /// <inheritdoc />
    public override string Operation => "edit";

    /// <summary>
    ///     Edits the specified cell with a new value, formula, or clears it.
    /// </summary>
    /// <param name="context">The workbook context.</param>
    /// <param name="parameters">
    ///     Required: cell (cell reference like "A1")
    ///     Optional: sheetIndex (default: 0), value, formula, clearValue (default: false)
    /// </param>
    /// <returns>Success message.</returns>
    public override string Execute(OperationContext<Workbook> context, OperationParameters parameters)
    {
        var cell = parameters.GetRequired<string>("cell");
        var sheetIndex = parameters.GetOptional("sheetIndex", 0);
        var value = parameters.GetOptional<string?>("value");
        var formula = parameters.GetOptional<string?>("formula");
        var clearValue = parameters.GetOptional("clearValue", false);

        ExcelCellHelper.ValidateCellAddress(cell);

        var workbook = context.Document;
        var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);
        var cellObj = worksheet.Cells[cell];

        if (clearValue)
            cellObj.PutValue("");
        else if (!string.IsNullOrEmpty(formula))
            cellObj.Formula = formula;
        else if (!string.IsNullOrEmpty(value))
            ExcelHelper.SetCellValue(cellObj, value);
        else
            throw new ArgumentException("Either value, formula, or clearValue must be provided");

        MarkModified(context);

        return Success($"Cell {cell} edited in sheet {sheetIndex}.");
    }
}
