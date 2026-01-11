using Aspose.Cells;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Helpers;

namespace AsposeMcpServer.Handlers.Excel.Group;

/// <summary>
///     Handler for ungrouping columns in Excel worksheets.
/// </summary>
public class UngroupExcelColumnsHandler : OperationHandlerBase<Workbook>
{
    /// <inheritdoc />
    public override string Operation => "ungroup_columns";

    /// <summary>
    ///     Ungroups columns.
    /// </summary>
    /// <param name="context">The workbook context.</param>
    /// <param name="parameters">
    ///     Required: startColumn, endColumn
    ///     Optional: sheetIndex (default: 0)
    /// </param>
    /// <returns>Success message with ungroup details.</returns>
    public override string Execute(OperationContext<Workbook> context, OperationParameters parameters)
    {
        ExcelGroupHelper.ValidateRequiredParams(Operation, parameters, "startColumn", "endColumn");

        var sheetIndex = parameters.GetOptional("sheetIndex", 0);
        var startColumn = parameters.GetRequired<int>("startColumn");
        var endColumn = parameters.GetRequired<int>("endColumn");

        ExcelGroupHelper.ValidateColumnRange(startColumn, endColumn);

        var workbook = context.Document;
        var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);
        worksheet.Cells.UngroupColumns(startColumn, endColumn);

        MarkModified(context);

        return Success($"Columns {startColumn}-{endColumn} ungrouped in sheet {sheetIndex}.");
    }
}
