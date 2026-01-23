using Aspose.Cells;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Helpers.Excel;
using AsposeMcpServer.Results.Common;

namespace AsposeMcpServer.Handlers.Excel.RowColumn;

/// <summary>
///     Handler for deleting columns from Excel worksheets.
/// </summary>
[ResultType(typeof(SuccessResult))]
public class DeleteColumnHandler : OperationHandlerBase<Workbook>
{
    /// <inheritdoc />
    public override string Operation => "delete_column";

    /// <summary>
    ///     Deletes columns starting from the specified position.
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="parameters">
    ///     Required: columnIndex
    ///     Optional: sheetIndex (default: 0), count (default: 1)
    /// </param>
    /// <returns>Success message with deletion details.</returns>
    public override object Execute(OperationContext<Workbook> context, OperationParameters parameters)
    {
        var p = ExtractDeleteColumnParameters(parameters);

        var workbook = context.Document;
        var worksheet = ExcelHelper.GetWorksheet(workbook, p.SheetIndex);

        worksheet.Cells.DeleteColumns(p.ColumnIndex, p.Count, true);

        MarkModified(context);

        return new SuccessResult { Message = $"Deleted {p.Count} column(s) starting from column {p.ColumnIndex}." };
    }

    private static DeleteColumnParameters ExtractDeleteColumnParameters(OperationParameters parameters)
    {
        var sheetIndex = parameters.GetOptional("sheetIndex", 0);
        var columnIndex = parameters.GetRequired<int>("columnIndex");
        var count = parameters.GetOptional("count", 1);

        return new DeleteColumnParameters(sheetIndex, columnIndex, count);
    }

    private sealed record DeleteColumnParameters(int SheetIndex, int ColumnIndex, int Count);
}
