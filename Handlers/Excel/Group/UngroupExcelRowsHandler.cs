using Aspose.Cells;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Helpers;

namespace AsposeMcpServer.Handlers.Excel.Group;

/// <summary>
///     Handler for ungrouping rows in Excel worksheets.
/// </summary>
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
    public override string Execute(OperationContext<Workbook> context, OperationParameters parameters)
    {
        ExcelGroupHelper.ValidateRequiredParams(Operation, parameters, "startRow", "endRow");

        var sheetIndex = parameters.GetOptional("sheetIndex", 0);
        var startRow = parameters.GetRequired<int>("startRow");
        var endRow = parameters.GetRequired<int>("endRow");

        ExcelGroupHelper.ValidateRowRange(startRow, endRow);

        var workbook = context.Document;
        var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);
        worksheet.Cells.UngroupRows(startRow, endRow);

        MarkModified(context);

        return Success($"Rows {startRow}-{endRow} ungrouped in sheet {sheetIndex}.");
    }
}
