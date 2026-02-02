using Aspose.Cells;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Helpers.Excel;
using AsposeMcpServer.Results.Common;

namespace AsposeMcpServer.Handlers.Excel.RowColumn;

/// <summary>
///     Handler for inserting rows into Excel worksheets.
/// </summary>
[ResultType(typeof(SuccessResult))]
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
    public override object Execute(OperationContext<Workbook> context, OperationParameters parameters)
    {
        var p = ExtractInsertRowParameters(parameters);

        var workbook = context.Document;
        var worksheet = ExcelHelper.GetWorksheet(workbook, p.SheetIndex);

        worksheet.Cells.InsertRows(p.RowIndex, p.Count);

        MarkModified(context);

        return new SuccessResult { Message = $"Inserted {p.Count} row(s) at row {p.RowIndex}." };
    }

    private static InsertRowParameters ExtractInsertRowParameters(OperationParameters parameters)
    {
        var sheetIndex = parameters.GetOptional("sheetIndex", 0);
        var rowIndex = parameters.GetRequired<int>("rowIndex");
        var count = parameters.GetOptional("count", 1);

        if (count <= 0)
            throw new ArgumentException($"Count must be greater than 0, got {count}");

        return new InsertRowParameters(sheetIndex, rowIndex, count);
    }

    private sealed record InsertRowParameters(int SheetIndex, int RowIndex, int Count);
}
