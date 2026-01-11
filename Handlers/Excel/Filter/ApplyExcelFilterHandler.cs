using Aspose.Cells;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Helpers;

namespace AsposeMcpServer.Handlers.Excel.Filter;

/// <summary>
///     Handler for applying auto filter to Excel worksheets.
/// </summary>
public class ApplyExcelFilterHandler : OperationHandlerBase<Workbook>
{
    /// <inheritdoc />
    public override string Operation => "apply";

    /// <summary>
    ///     Applies auto filter dropdown buttons to a range.
    /// </summary>
    /// <param name="context">The workbook context.</param>
    /// <param name="parameters">
    ///     Required: range
    ///     Optional: sheetIndex (default: 0)
    /// </param>
    /// <returns>Success message with filter details.</returns>
    public override string Execute(OperationContext<Workbook> context, OperationParameters parameters)
    {
        var sheetIndex = parameters.GetOptional("sheetIndex", 0);
        var range = parameters.GetRequired<string>("range");

        var workbook = context.Document;
        var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);

        ExcelHelper.CreateRange(worksheet.Cells, range);
        worksheet.AutoFilter.Range = range;

        MarkModified(context);

        return Success($"Auto filter applied to range {range} in sheet {sheetIndex}.");
    }
}
