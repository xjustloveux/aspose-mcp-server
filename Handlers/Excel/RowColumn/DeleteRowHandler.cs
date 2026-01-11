using Aspose.Cells;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Helpers;

namespace AsposeMcpServer.Handlers.Excel.RowColumn;

/// <summary>
///     Handler for deleting rows from Excel worksheets.
/// </summary>
public class DeleteRowHandler : OperationHandlerBase<Workbook>
{
    /// <inheritdoc />
    public override string Operation => "delete_row";

    /// <summary>
    ///     Deletes rows starting from the specified position.
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="parameters">
    ///     Required: rowIndex
    ///     Optional: sheetIndex (default: 0), count (default: 1)
    /// </param>
    /// <returns>Success message with deletion details.</returns>
    public override string Execute(OperationContext<Workbook> context, OperationParameters parameters)
    {
        var sheetIndex = parameters.GetOptional("sheetIndex", 0);
        var rowIndex = parameters.GetRequired<int>("rowIndex");
        var count = parameters.GetOptional("count", 1);

        var workbook = context.Document;
        var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);

        worksheet.Cells.DeleteRows(rowIndex, count);

        MarkModified(context);

        return Success($"Deleted {count} row(s) starting from row {rowIndex}.");
    }
}
