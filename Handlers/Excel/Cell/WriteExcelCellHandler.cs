using Aspose.Cells;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Helpers;

namespace AsposeMcpServer.Handlers.Excel.Cell;

/// <summary>
///     Handler for writing values to Excel cells.
/// </summary>
public class WriteExcelCellHandler : OperationHandlerBase<Workbook>
{
    /// <inheritdoc />
    public override string Operation => "write";

    /// <summary>
    ///     Writes a value to the specified cell.
    /// </summary>
    /// <param name="context">The workbook context.</param>
    /// <param name="parameters">
    ///     Required: cell (cell reference like "A1"), value
    ///     Optional: sheetIndex (default: 0)
    /// </param>
    /// <returns>Success message.</returns>
    public override string Execute(OperationContext<Workbook> context, OperationParameters parameters)
    {
        var cell = parameters.GetRequired<string>("cell");
        var value = parameters.GetRequired<string>("value");
        var sheetIndex = parameters.GetOptional("sheetIndex", 0);

        if (string.IsNullOrEmpty(value))
            throw new ArgumentException("value is required for write operation");

        ExcelCellHelper.ValidateCellAddress(cell);

        var workbook = context.Document;
        var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);
        var cellObj = worksheet.Cells[cell];

        ExcelHelper.SetCellValue(cellObj, value);

        MarkModified(context);

        return Success($"Cell {cell} written with value '{value}' in sheet {sheetIndex}.");
    }
}
