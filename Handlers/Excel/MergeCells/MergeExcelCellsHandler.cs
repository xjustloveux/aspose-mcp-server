using Aspose.Cells;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Helpers.Excel;
using AsposeMcpServer.Results.Common;

namespace AsposeMcpServer.Handlers.Excel.MergeCells;

/// <summary>
///     Handler for merging cells in Excel workbooks.
/// </summary>
[ResultType(typeof(SuccessResult))]
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
    public override object Execute(OperationContext<Workbook> context, OperationParameters parameters)
    {
        var p = ExtractMergeCellsParameters(parameters);

        var workbook = context.Document;
        var worksheet = ExcelHelper.GetWorksheet(workbook, p.SheetIndex);
        var cells = worksheet.Cells;

        var cellRange = ExcelHelper.CreateRange(cells, p.Range);

        if (cellRange is { RowCount: 1, ColumnCount: 1 })
            throw new ArgumentException(
                $"Cannot merge a single cell. Range '{p.Range}' must include at least 2 cells.");

        cellRange.Merge();

        MarkModified(context);

        return new SuccessResult
        {
            Message =
                $"Range {p.Range} merged ({cellRange.RowCount} rows x {cellRange.ColumnCount} columns)."
        };
    }

    private static MergeCellsParameters ExtractMergeCellsParameters(OperationParameters parameters)
    {
        return new MergeCellsParameters(
            parameters.GetRequired<string>("range"),
            parameters.GetOptional("sheetIndex", 0)
        );
    }

    private sealed record MergeCellsParameters(string Range, int SheetIndex);
}
