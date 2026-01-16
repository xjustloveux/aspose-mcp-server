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

        var p = ExtractUngroupColumnsParameters(parameters);

        ExcelGroupHelper.ValidateColumnRange(p.StartColumn, p.EndColumn);

        var workbook = context.Document;
        var worksheet = ExcelHelper.GetWorksheet(workbook, p.SheetIndex);
        worksheet.Cells.UngroupColumns(p.StartColumn, p.EndColumn);

        MarkModified(context);

        return Success($"Columns {p.StartColumn}-{p.EndColumn} ungrouped in sheet {p.SheetIndex}.");
    }

    private static UngroupColumnsParameters ExtractUngroupColumnsParameters(OperationParameters parameters)
    {
        var sheetIndex = parameters.GetOptional("sheetIndex", 0);
        var startColumn = parameters.GetRequired<int>("startColumn");
        var endColumn = parameters.GetRequired<int>("endColumn");

        return new UngroupColumnsParameters(sheetIndex, startColumn, endColumn);
    }

    private sealed record UngroupColumnsParameters(int SheetIndex, int StartColumn, int EndColumn);
}
