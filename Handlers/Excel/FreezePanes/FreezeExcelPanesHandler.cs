using Aspose.Cells;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Helpers.Excel;
using AsposeMcpServer.Results.Common;

namespace AsposeMcpServer.Handlers.Excel.FreezePanes;

/// <summary>
///     Handler for freezing panes in Excel worksheets.
/// </summary>
[ResultType(typeof(SuccessResult))]
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
    public override object Execute(OperationContext<Workbook> context, OperationParameters parameters)
    {
        var p = ExtractFreezeParameters(parameters);

        var workbook = context.Document;
        var worksheet = ExcelHelper.GetWorksheet(workbook, p.SheetIndex);

        // FreezePanes(row, column, freezedRows, freezedColumns): the first pair is the split cell,
        // the second is how many rows and columns the frozen pane holds. Freezing N rows means the
        // split sits at row N and the pane holds N rows, so both are the same number. Passing N+1
        // as the split put it one row too far, which opened the sheet scrolled a row down and made
        // the value read back differ from the one excel_view_settings writes for the same request.
        worksheet.FreezePanes(p.Row, p.Column, p.Row, p.Column);

        MarkModified(context);

        return new SuccessResult { Message = $"Frozen panes at row {p.Row}, column {p.Column}." };
    }

    private static FreezeParameters ExtractFreezeParameters(OperationParameters parameters)
    {
        var sheetIndex = parameters.GetOptional("sheetIndex", 0);
        var row = parameters.GetRequired<int>("row");
        var column = parameters.GetRequired<int>("column");

        return new FreezeParameters(sheetIndex, row, column);
    }

    private sealed record FreezeParameters(int SheetIndex, int Row, int Column);
}
