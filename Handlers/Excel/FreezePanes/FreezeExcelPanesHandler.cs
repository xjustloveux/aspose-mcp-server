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

        worksheet.FreezePanes(p.Row + 1, p.Column + 1, p.Row, p.Column);

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
