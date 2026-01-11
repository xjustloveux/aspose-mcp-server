using Aspose.Cells;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Helpers;

namespace AsposeMcpServer.Handlers.Excel.MergeCells;

/// <summary>
///     Handler for merging cells in Excel workbooks.
/// </summary>
public class MergeExcelCellsHandler : OperationHandlerBase<Workbook>
{
    /// <inheritdoc />
    public override string Operation => "merge";

    /// <summary>
    ///     Merges cells in a range.
    /// </summary>
    /// <param name="context">The workbook context.</param>
    /// <param name="parameters">
    ///     Required: range (cell range to merge, e.g., "A1:C3")
    ///     Optional: sheetIndex (0-based, default: 0)
    /// </param>
    /// <returns>Success message with merge details.</returns>
    public override string Execute(OperationContext<Workbook> context, OperationParameters parameters)
    {
        var range = parameters.GetRequired<string>("range");
        var sheetIndex = parameters.GetOptional("sheetIndex", 0);

        var workbook = context.Document;
        var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);
        var cells = worksheet.Cells;

        var cellRange = ExcelHelper.CreateRange(cells, range);

        if (cellRange is { RowCount: 1, ColumnCount: 1 })
            throw new ArgumentException(
                $"Cannot merge a single cell. Range '{range}' must include at least 2 cells.");

        cellRange.Merge();

        MarkModified(context);

        return Success(
            $"Range {range} merged ({cellRange.RowCount} rows x {cellRange.ColumnCount} columns).");
    }
}
