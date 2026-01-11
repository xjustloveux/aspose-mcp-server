using Aspose.Cells;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Helpers;

namespace AsposeMcpServer.Handlers.Excel.Group;

/// <summary>
///     Handler for grouping columns in Excel worksheets.
/// </summary>
public class GroupExcelColumnsHandler : OperationHandlerBase<Workbook>
{
    /// <inheritdoc />
    public override string Operation => "group_columns";

    /// <summary>
    ///     Groups columns together.
    /// </summary>
    /// <param name="context">The workbook context.</param>
    /// <param name="parameters">
    ///     Required: startColumn, endColumn
    ///     Optional: sheetIndex (default: 0), isCollapsed (default: false)
    /// </param>
    /// <returns>Success message with group details.</returns>
    public override string Execute(OperationContext<Workbook> context, OperationParameters parameters)
    {
        ExcelGroupHelper.ValidateRequiredParams(Operation, parameters, "startColumn", "endColumn");

        var sheetIndex = parameters.GetOptional("sheetIndex", 0);
        var startColumn = parameters.GetRequired<int>("startColumn");
        var endColumn = parameters.GetRequired<int>("endColumn");
        var isCollapsed = parameters.GetOptional("isCollapsed", false);

        ExcelGroupHelper.ValidateColumnRange(startColumn, endColumn);

        var workbook = context.Document;
        var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);
        worksheet.Cells.GroupColumns(startColumn, endColumn, isCollapsed);

        MarkModified(context);

        return Success($"Columns {startColumn}-{endColumn} grouped in sheet {sheetIndex}.");
    }
}
