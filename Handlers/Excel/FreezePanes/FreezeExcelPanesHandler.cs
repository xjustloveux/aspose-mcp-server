using Aspose.Cells;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Helpers;

namespace AsposeMcpServer.Handlers.Excel.FreezePanes;

/// <summary>
///     Handler for freezing panes in Excel worksheets.
/// </summary>
public class FreezeExcelPanesHandler : OperationHandlerBase<Workbook>
{
    /// <inheritdoc />
    public override string Operation => "freeze";

    /// <summary>
    ///     Freezes panes at the specified row and column.
    /// </summary>
    /// <param name="context">The workbook context.</param>
    /// <param name="parameters">
    ///     Required: row, column
    ///     Optional: sheetIndex (default: 0)
    /// </param>
    /// <returns>Success message with freeze details.</returns>
    public override string Execute(OperationContext<Workbook> context, OperationParameters parameters)
    {
        var sheetIndex = parameters.GetOptional("sheetIndex", 0);
        var row = parameters.GetRequired<int>("row");
        var column = parameters.GetRequired<int>("column");

        var workbook = context.Document;
        var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);

        worksheet.FreezePanes(row + 1, column + 1, row, column);

        MarkModified(context);

        return Success($"Frozen panes at row {row}, column {column}.");
    }
}
