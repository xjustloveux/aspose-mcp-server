using Aspose.Cells;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Helpers.Excel;
using AsposeMcpServer.Results.Common;

namespace AsposeMcpServer.Handlers.Excel.RowColumn;

/// <summary>
///     Handler for inserting columns into Excel worksheets.
/// </summary>
[ResultType(typeof(SuccessResult))]
public class InsertColumnHandler : OperationHandlerBase<Workbook>
{
    /// <inheritdoc />
    public override string Operation => "insert_column";

    /// <summary>
    ///     Inserts columns at the specified position.
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="parameters">
    ///     Required: columnIndex
    ///     Optional: sheetIndex (default: 0), count (default: 1)
    /// </param>
    /// <returns>Success message with insertion details.</returns>
    public override object Execute(OperationContext<Workbook> context, OperationParameters parameters)
    {
        var p = ExtractInsertColumnParameters(parameters);

        var workbook = context.Document;
        var worksheet = ExcelHelper.GetWorksheet(workbook, p.SheetIndex);

        worksheet.Cells.InsertColumns(p.ColumnIndex, p.Count);

        MarkModified(context);

        return new SuccessResult { Message = $"Inserted {p.Count} column(s) at column {p.ColumnIndex}." };
    }

    private static InsertColumnParameters ExtractInsertColumnParameters(OperationParameters parameters)
    {
        var sheetIndex = parameters.GetOptional("sheetIndex", 0);
        var columnIndex = parameters.GetRequired<int>("columnIndex");
        var count = parameters.GetOptional("count", 1);

        return new InsertColumnParameters(sheetIndex, columnIndex, count);
    }

    private sealed record InsertColumnParameters(int SheetIndex, int ColumnIndex, int Count);
}
