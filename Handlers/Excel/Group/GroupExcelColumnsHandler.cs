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

        var p = ExtractGroupColumnsParameters(parameters);

        ExcelGroupHelper.ValidateColumnRange(p.StartColumn, p.EndColumn);

        var workbook = context.Document;
        var worksheet = ExcelHelper.GetWorksheet(workbook, p.SheetIndex);
        worksheet.Cells.GroupColumns(p.StartColumn, p.EndColumn, p.IsCollapsed);

        MarkModified(context);

        return Success($"Columns {p.StartColumn}-{p.EndColumn} grouped in sheet {p.SheetIndex}.");
    }

    private static GroupColumnsParameters ExtractGroupColumnsParameters(OperationParameters parameters)
    {
        var sheetIndex = parameters.GetOptional("sheetIndex", 0);
        var startColumn = parameters.GetRequired<int>("startColumn");
        var endColumn = parameters.GetRequired<int>("endColumn");
        var isCollapsed = parameters.GetOptional("isCollapsed", false);

        return new GroupColumnsParameters(sheetIndex, startColumn, endColumn, isCollapsed);
    }

    private record GroupColumnsParameters(int SheetIndex, int StartColumn, int EndColumn, bool IsCollapsed);
}
