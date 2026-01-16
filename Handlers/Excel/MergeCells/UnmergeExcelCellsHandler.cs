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
        var p = ExtractUnmergeCellsParameters(parameters);

        var workbook = context.Document;
        var worksheet = ExcelHelper.GetWorksheet(workbook, p.SheetIndex);
        var cells = worksheet.Cells;

        var cellRange = ExcelHelper.CreateRange(cells, p.Range);

        cellRange.UnMerge();

        MarkModified(context);

        return Success($"Range {p.Range} unmerged.");
    }

    private static UnmergeCellsParameters ExtractUnmergeCellsParameters(OperationParameters parameters)
    {
        return new UnmergeCellsParameters(
            parameters.GetRequired<string>("range"),
            parameters.GetOptional("sheetIndex", 0)
        );
    }

    private sealed record UnmergeCellsParameters(string Range, int SheetIndex);
}
