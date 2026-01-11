using Aspose.Cells;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Helpers;

namespace AsposeMcpServer.Handlers.Excel.RowColumn;

/// <summary>
///     Handler for inserting rows into Excel worksheets.
/// </summary>
public class InsertRowHandler : OperationHandlerBase<Workbook>
{
    /// <inheritdoc />
    public override string Operation => "insert_row";

    /// <summary>
    ///     Inserts rows at the specified position.
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="parameters">
    ///     Required: rowIndex
    ///     Optional: sheetIndex (default: 0), count (default: 1)
    /// </param>
    /// <returns>Success message with insertion details.</returns>
    public override string Execute(OperationContext<Workbook> context, OperationParameters parameters)
    {
        var sheetIndex = parameters.GetOptional("sheetIndex", 0);
        var rowIndex = parameters.GetRequired<int>("rowIndex");
        var count = parameters.GetOptional("count", 1);

        var workbook = context.Document;
        var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);

        worksheet.Cells.InsertRows(rowIndex, count);

        MarkModified(context);

        return Success($"Inserted {count} row(s) at row {rowIndex}.");
    }
}
