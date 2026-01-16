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
        var p = ExtractDeleteRowParameters(parameters);

        var workbook = context.Document;
        var worksheet = ExcelHelper.GetWorksheet(workbook, p.SheetIndex);

        worksheet.Cells.DeleteRows(p.RowIndex, p.Count);

        MarkModified(context);

        return Success($"Deleted {p.Count} row(s) starting from row {p.RowIndex}.");
    }

    private static DeleteRowParameters ExtractDeleteRowParameters(OperationParameters parameters)
    {
        var sheetIndex = parameters.GetOptional("sheetIndex", 0);
        var rowIndex = parameters.GetRequired<int>("rowIndex");
        var count = parameters.GetOptional("count", 1);

        return new DeleteRowParameters(sheetIndex, rowIndex, count);
    }

    private sealed record DeleteRowParameters(int SheetIndex, int RowIndex, int Count);
}
