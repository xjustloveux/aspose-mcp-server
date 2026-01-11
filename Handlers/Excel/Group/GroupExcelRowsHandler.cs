using Aspose.Cells;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Helpers;

namespace AsposeMcpServer.Handlers.Excel.Group;

/// <summary>
///     Handler for grouping rows in Excel worksheets.
/// </summary>
public class GroupExcelRowsHandler : OperationHandlerBase<Workbook>
{
    /// <inheritdoc />
    public override string Operation => "group_rows";

    /// <summary>
    ///     Groups rows together.
    /// </summary>
    /// <param name="context">The workbook context.</param>
    /// <param name="parameters">
    ///     Required: startRow, endRow
    ///     Optional: sheetIndex (default: 0), isCollapsed (default: false)
    /// </param>
    /// <returns>Success message with group details.</returns>
    public override string Execute(OperationContext<Workbook> context, OperationParameters parameters)
    {
        ExcelGroupHelper.ValidateRequiredParams(Operation, parameters, "startRow", "endRow");

        var sheetIndex = parameters.GetOptional("sheetIndex", 0);
        var startRow = parameters.GetRequired<int>("startRow");
        var endRow = parameters.GetRequired<int>("endRow");
        var isCollapsed = parameters.GetOptional("isCollapsed", false);

        ExcelGroupHelper.ValidateRowRange(startRow, endRow);

        var workbook = context.Document;
        var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);
        worksheet.Cells.GroupRows(startRow, endRow, isCollapsed);

        MarkModified(context);

        return Success($"Rows {startRow}-{endRow} grouped in sheet {sheetIndex}.");
    }
}
