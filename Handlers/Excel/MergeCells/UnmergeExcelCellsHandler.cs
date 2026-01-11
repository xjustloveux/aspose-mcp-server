using Aspose.Cells;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Helpers;

namespace AsposeMcpServer.Handlers.Excel.MergeCells;

/// <summary>
///     Handler for unmerging cells in Excel workbooks.
/// </summary>
public class UnmergeExcelCellsHandler : OperationHandlerBase<Workbook>
{
    /// <inheritdoc />
    public override string Operation => "unmerge";

    /// <summary>
    ///     Unmerges cells in a range.
    /// </summary>
    /// <param name="context">The workbook context.</param>
    /// <param name="parameters">
    ///     Required: range (cell range to unmerge, e.g., "A1:C3")
    ///     Optional: sheetIndex (0-based, default: 0)
    /// </param>
    /// <returns>Success message with unmerge details.</returns>
    public override string Execute(OperationContext<Workbook> context, OperationParameters parameters)
    {
        var range = parameters.GetRequired<string>("range");
        var sheetIndex = parameters.GetOptional("sheetIndex", 0);

        var workbook = context.Document;
        var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);
        var cells = worksheet.Cells;

        var cellRange = ExcelHelper.CreateRange(cells, range);

        cellRange.UnMerge();

        MarkModified(context);

        return Success($"Range {range} unmerged.");
    }
}
