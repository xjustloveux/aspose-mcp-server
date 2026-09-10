using Aspose.Cells;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Errors.Excel;
using AsposeMcpServer.Helpers.Excel;
using AsposeMcpServer.Results.Common;

namespace AsposeMcpServer.Handlers.Excel.DataOperations;

/// <summary>
///     Handler for sorting data in Excel worksheets.
/// </summary>
[ResultType(typeof(SuccessResult))]
public class SortDataHandler : OperationHandlerBase<Workbook>
{
    /// <inheritdoc />
    public override string Operation => "sort";

    /// <summary>
    ///     Sorts data in a specified range.
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="parameters">
    ///     Required: range
    ///     Optional: sheetIndex, sortColumn, ascending, hasHeader
    /// </param>
    /// <returns>Success message with sort details.</returns>
    public override object Execute(OperationContext<Workbook> context, OperationParameters parameters)
    {
        var sortParams = ExtractSortDataParameters(parameters);

        if (string.IsNullOrEmpty(sortParams.Range))
            throw new ArgumentException("range is required for sort operation");

        try
        {
            var workbook = context.Document;
            var worksheet = ExcelHelper.GetWorksheet(workbook, sortParams.SheetIndex);
            var cells = worksheet.Cells;
            var cellRange = ExcelHelper.CreateRange(cells, sortParams.Range);

            if (sortParams.SortColumn < 0 || sortParams.SortColumn >= cellRange.ColumnCount)
                throw new ArgumentException(
                    $"sortColumn must be between 0 and {cellRange.ColumnCount - 1} for range {sortParams.Range}");

            // Aspose's own sorter relocates whole rows, so formulas, styles, comments, hyperlinks
            // and validation travel with their data. Reading Cell.Value into a list and writing it
            // back only moved values: a formula became the constant it happened to evaluate to, and
            // every other per-cell attribute stayed behind at the old coordinates.
            var sorter = workbook.DataSorter;
            sorter.Clear();
            sorter.HasHeaders = sortParams.HasHeader;
            sorter.AddKey(cellRange.FirstColumn + sortParams.SortColumn,
                sortParams.Ascending ? SortOrder.Ascending : SortOrder.Descending);
            sorter.Sort(cells, new CellArea
            {
                StartRow = cellRange.FirstRow,
                StartColumn = cellRange.FirstColumn,
                EndRow = cellRange.FirstRow + cellRange.RowCount - 1,
                EndColumn = cellRange.FirstColumn + cellRange.ColumnCount - 1
            });
            sorter.Clear();

            MarkModified(context);

            return new SuccessResult
            {
                Message =
                    $"Sorted range {sortParams.Range} by column {sortParams.SortColumn} ({(sortParams.Ascending ? "ascending" : "descending")})."
            };
        }
        catch (CellsException ex)
        {
            throw CellsErrorTranslator.Translate(ex);
        }
    }

    /// <summary>
    ///     Extracts sort data parameters from the operation parameters.
    /// </summary>
    /// <param name="parameters">The operation parameters.</param>
    /// <returns>The extracted sort data parameters.</returns>
    private static SortDataParameters ExtractSortDataParameters(OperationParameters parameters)
    {
        return new SortDataParameters(
            parameters.GetOptional<string?>("range"),
            parameters.GetOptional("sheetIndex", 0),
            parameters.GetOptional("sortColumn", 0),
            parameters.GetOptional("ascending", true),
            parameters.GetOptional("hasHeader", false)
        );
    }

    /// <summary>
    ///     Parameters for sort data operation.
    /// </summary>
    /// <param name="Range">The cell range to sort.</param>
    /// <param name="SheetIndex">The worksheet index (0-based).</param>
    /// <param name="SortColumn">The column index to sort by (0-based within the range).</param>
    /// <param name="Ascending">Whether to sort in ascending order.</param>
    /// <param name="HasHeader">Whether the range has a header row that should not be sorted.</param>
    private sealed record SortDataParameters(
        string? Range,
        int SheetIndex,
        int SortColumn,
        bool Ascending,
        bool HasHeader);
}
