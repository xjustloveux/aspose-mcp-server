using Aspose.Cells;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Helpers;

namespace AsposeMcpServer.Handlers.Excel.Cell;

/// <summary>
///     Handler for clearing Excel cell content and/or format.
/// </summary>
public class ClearExcelCellHandler : OperationHandlerBase<Workbook>
{
    /// <inheritdoc />
    public override string Operation => "clear";

    /// <summary>
    ///     Clears the content and/or format of the specified cell.
    /// </summary>
    /// <param name="context">The workbook context.</param>
    /// <param name="parameters">
    ///     Required: cell (cell reference like "A1")
    ///     Optional: sheetIndex (default: 0), clearContent (default: true), clearFormat (default: false)
    /// </param>
    /// <returns>Success message.</returns>
    public override string Execute(OperationContext<Workbook> context, OperationParameters parameters)
    {
        var cell = parameters.GetRequired<string>("cell");
        var sheetIndex = parameters.GetOptional("sheetIndex", 0);
        var clearContent = parameters.GetOptional("clearContent", true);
        var clearFormat = parameters.GetOptional("clearFormat", false);

        ExcelCellHelper.ValidateCellAddress(cell);

        var workbook = context.Document;
        var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);
        var cellObj = worksheet.Cells[cell];

        if (clearContent && clearFormat)
        {
            cellObj.PutValue("");
            var defaultStyle = workbook.CreateStyle();
            cellObj.SetStyle(defaultStyle);
        }
        else if (clearContent)
        {
            cellObj.PutValue("");
        }
        else if (clearFormat)
        {
            var defaultStyle = workbook.CreateStyle();
            cellObj.SetStyle(defaultStyle);
        }

        MarkModified(context);

        return Success($"Cell {cell} cleared in sheet {sheetIndex}.");
    }
}
