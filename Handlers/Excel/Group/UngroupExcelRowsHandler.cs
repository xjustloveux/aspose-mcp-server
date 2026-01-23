using Aspose.Cells;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Helpers.Excel;
using AsposeMcpServer.Results.Common;

namespace AsposeMcpServer.Handlers.Excel.Group;

/// <summary>
///     Handler for ungrouping rows in Excel worksheets.
/// </summary>
[ResultType(typeof(SuccessResult))]
public class UngroupExcelRowsHandler : OperationHandlerBase<Workbook>
{
    /// <inheritdoc />
    public override string Operation => "ungroup_rows";

    /// <summary>
    ///     Ungroups rows.
    /// </summary>
    /// <param name="context">The workbook context.</param>
    /// <param name="parameters">
    ///     Required: startRow, endRow
    ///     Optional: sheetIndex (default: 0)
    /// </param>
    /// <returns>Success message with ungroup details.</returns>
    public override object Execute(OperationContext<Workbook> context, OperationParameters parameters)
    {
        ExcelGroupHelper.ValidateRequiredParams(Operation, parameters, "startRow", "endRow");

        var p = ExtractUngroupRowsParameters(parameters);

        ExcelGroupHelper.ValidateRowRange(p.StartRow, p.EndRow);

        var workbook = context.Document;
        var worksheet = ExcelHelper.GetWorksheet(workbook, p.SheetIndex);
        worksheet.Cells.UngroupRows(p.StartRow, p.EndRow);

        MarkModified(context);

        return new SuccessResult { Message = $"Rows {p.StartRow}-{p.EndRow} ungrouped in sheet {p.SheetIndex}." };
    }

    private static UngroupRowsParameters ExtractUngroupRowsParameters(OperationParameters parameters)
    {
        var sheetIndex = parameters.GetOptional("sheetIndex", 0);
        var startRow = parameters.GetRequired<int>("startRow");
        var endRow = parameters.GetRequired<int>("endRow");

        return new UngroupRowsParameters(sheetIndex, startRow, endRow);
    }

    private sealed record UngroupRowsParameters(int SheetIndex, int StartRow, int EndRow);
}
