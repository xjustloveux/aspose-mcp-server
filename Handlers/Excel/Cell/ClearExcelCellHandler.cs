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
    ///     Required: cell
    ///     Optional: sheetIndex, clearContent, clearFormat
    /// </param>
    /// <returns>Success message.</returns>
    public override string Execute(OperationContext<Workbook> context, OperationParameters parameters)
    {
        var clearParams = ExtractClearParameters(parameters);

        ExcelCellHelper.ValidateCellAddress(clearParams.Cell);

        var workbook = context.Document;
        var worksheet = ExcelHelper.GetWorksheet(workbook, clearParams.SheetIndex);
        var cellObj = worksheet.Cells[clearParams.Cell];

        if (clearParams is { ClearContent: true, ClearFormat: true })
        {
            cellObj.PutValue("");
            var defaultStyle = workbook.CreateStyle();
            cellObj.SetStyle(defaultStyle);
        }
        else if (clearParams.ClearContent)
        {
            cellObj.PutValue("");
        }
        else if (clearParams.ClearFormat)
        {
            var defaultStyle = workbook.CreateStyle();
            cellObj.SetStyle(defaultStyle);
        }

        MarkModified(context);

        return Success($"Cell {clearParams.Cell} cleared in sheet {clearParams.SheetIndex}.");
    }

    /// <summary>
    ///     Extracts clear parameters from operation parameters.
    /// </summary>
    /// <param name="parameters">The operation parameters.</param>
    /// <returns>The extracted clear parameters.</returns>
    private static ClearParameters ExtractClearParameters(OperationParameters parameters)
    {
        return new ClearParameters(
            parameters.GetRequired<string>("cell"),
            parameters.GetOptional("sheetIndex", 0),
            parameters.GetOptional("clearContent", true),
            parameters.GetOptional("clearFormat", false)
        );
    }

    /// <summary>
    ///     Record to hold clear cell parameters.
    /// </summary>
    private record ClearParameters(string Cell, int SheetIndex, bool ClearContent, bool ClearFormat);
}
