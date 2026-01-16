using Aspose.Cells;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Helpers;

namespace AsposeMcpServer.Handlers.Excel.DataOperations;

/// <summary>
///     Handler for sorting data in Excel worksheets.
/// </summary>
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
    public override string Execute(OperationContext<Workbook> context, OperationParameters parameters)
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

            var rows = ExtractRows(cells, cellRange, sortParams.HasHeader);
            var sortedRows = SortRows(rows, sortParams.SortColumn, sortParams.Ascending, sortParams.HasHeader);
            WriteRowsToSheet(cells, cellRange, sortedRows);

            MarkModified(context);

            return Success(
                $"Sorted range {sortParams.Range} by column {sortParams.SortColumn} ({(sortParams.Ascending ? "ascending" : "descending")}).");
        }
        catch (CellsException ex)
        {
            throw new ArgumentException($"Excel operation failed for range '{sortParams.Range}': {ex.Message}");
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
    ///     Extracts rows from a cell range.
    /// </summary>
    /// <param name="cells">The cells collection.</param>
    /// <param name="cellRange">The range to extract rows from.</param>
    /// <param name="hasHeader">Whether the range has a header row.</param>
    /// <returns>A list of rows, each containing cell values.</returns>
    private static List<List<object?>> ExtractRows(Cells cells, Aspose.Cells.Range cellRange, bool hasHeader)
    {
        List<List<object?>> rows = [];
        var startRow = hasHeader ? cellRange.FirstRow + 1 : cellRange.FirstRow;

        if (hasHeader) rows.Add(ExtractRow(cells, cellRange.FirstRow, cellRange.FirstColumn, cellRange.ColumnCount));

        for (var row = startRow; row < cellRange.FirstRow + cellRange.RowCount; row++)
            rows.Add(ExtractRow(cells, row, cellRange.FirstColumn, cellRange.ColumnCount));

        return rows;
    }

    /// <summary>
    ///     Extracts a single row of cell values.
    /// </summary>
    /// <param name="cells">The cells collection.</param>
    /// <param name="row">The row index.</param>
    /// <param name="startCol">The starting column index.</param>
    /// <param name="colCount">The number of columns to extract.</param>
    /// <returns>A list of cell values for the row.</returns>
    private static List<object?> ExtractRow(Cells cells, int row, int startCol, int colCount)
    {
        List<object?> rowData = [];
        for (var col = startCol; col < startCol + colCount; col++) rowData.Add(cells[row, col].Value);

        return rowData;
    }

    /// <summary>
    ///     Sorts the rows by the specified column.
    /// </summary>
    /// <param name="rows">The rows to sort.</param>
    /// <param name="sortColumn">The column index to sort by.</param>
    /// <param name="ascending">Whether to sort in ascending order.</param>
    /// <param name="hasHeader">Whether the first row is a header.</param>
    /// <returns>The sorted rows with header preserved if present.</returns>
    private static List<List<object?>> SortRows(List<List<object?>> rows, int sortColumn, bool ascending,
        bool hasHeader)
    {
        var dataRows = hasHeader ? rows.Skip(1).ToList() : rows;
        dataRows.Sort((a, b) => CompareRows(a, b, sortColumn, ascending));

        if (!hasHeader) return dataRows;

        List<List<object?>> result = [rows[0]];
        result.AddRange(dataRows);
        return result;
    }

    /// <summary>
    ///     Compares two rows by the specified column.
    /// </summary>
    /// <param name="a">The first row.</param>
    /// <param name="b">The second row.</param>
    /// <param name="sortColumn">The column index to compare.</param>
    /// <param name="ascending">Whether to sort in ascending order.</param>
    /// <returns>A comparison result for sorting.</returns>
    private static int CompareRows(List<object?> a, List<object?> b, int sortColumn, bool ascending)
    {
        var aVal = a[sortColumn];
        var bVal = b[sortColumn];

        if (aVal == null && bVal == null) return 0;
        if (aVal == null) return ascending ? -1 : 1;
        if (bVal == null) return ascending ? 1 : -1;

        var comparison = Comparer<object>.Default.Compare(aVal, bVal);
        return ascending ? comparison : -comparison;
    }

    /// <summary>
    ///     Writes the sorted rows back to the worksheet.
    /// </summary>
    /// <param name="cells">The cells collection.</param>
    /// <param name="cellRange">The target range.</param>
    /// <param name="rows">The rows to write.</param>
    private static void WriteRowsToSheet(Cells cells, Aspose.Cells.Range cellRange, List<List<object?>> rows)
    {
        for (var i = 0; i < rows.Count; i++)
        {
            var rowData = rows[i];
            var targetRow = cellRange.FirstRow + i;
            for (var j = 0; j < rowData.Count; j++) cells[targetRow, cellRange.FirstColumn + j].Value = rowData[j];
        }
    }

    /// <summary>
    ///     Parameters for sort data operation.
    /// </summary>
    /// <param name="Range">The cell range to sort.</param>
    /// <param name="SheetIndex">The worksheet index (0-based).</param>
    /// <param name="SortColumn">The column index to sort by (0-based within the range).</param>
    /// <param name="Ascending">Whether to sort in ascending order.</param>
    /// <param name="HasHeader">Whether the range has a header row that should not be sorted.</param>
    private record SortDataParameters(
        string? Range,
        int SheetIndex,
        int SortColumn,
        bool Ascending,
        bool HasHeader);
}
