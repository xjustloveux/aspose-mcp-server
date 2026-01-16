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
        var range = parameters.GetOptional<string?>("range");
        var sheetIndex = parameters.GetOptional("sheetIndex", 0);
        var sortColumn = parameters.GetOptional("sortColumn", 0);
        var ascending = parameters.GetOptional("ascending", true);
        var hasHeader = parameters.GetOptional("hasHeader", false);

        if (string.IsNullOrEmpty(range))
            throw new ArgumentException("range is required for sort operation");

        try
        {
            var workbook = context.Document;
            var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);
            var cells = worksheet.Cells;
            var cellRange = ExcelHelper.CreateRange(cells, range);

            var rows = ExtractRows(cells, cellRange, hasHeader);
            var sortedRows = SortRows(rows, sortColumn, ascending, hasHeader);
            WriteRowsToSheet(cells, cellRange, sortedRows);

            MarkModified(context);

            return Success(
                $"Sorted range {range} by column {sortColumn} ({(ascending ? "ascending" : "descending")}).");
        }
        catch (CellsException ex)
        {
            throw new ArgumentException($"Excel operation failed for range '{range}': {ex.Message}");
        }
    }

    private static List<List<object?>> ExtractRows(Cells cells, Aspose.Cells.Range cellRange, bool hasHeader)
    {
        List<List<object?>> rows = [];
        var startRow = hasHeader ? cellRange.FirstRow + 1 : cellRange.FirstRow;

        if (hasHeader) rows.Add(ExtractRow(cells, cellRange.FirstRow, cellRange.FirstColumn, cellRange.ColumnCount));

        for (var row = startRow; row < cellRange.FirstRow + cellRange.RowCount; row++)
            rows.Add(ExtractRow(cells, row, cellRange.FirstColumn, cellRange.ColumnCount));

        return rows;
    }

    private static List<object?> ExtractRow(Cells cells, int row, int startCol, int colCount)
    {
        List<object?> rowData = [];
        for (var col = startCol; col < startCol + colCount; col++) rowData.Add(cells[row, col].Value);

        return rowData;
    }

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

    private static void WriteRowsToSheet(Cells cells, Aspose.Cells.Range cellRange, List<List<object?>> rows)
    {
        for (var i = 0; i < rows.Count; i++)
        {
            var rowData = rows[i];
            var targetRow = cellRange.FirstRow + i;
            for (var j = 0; j < rowData.Count; j++) cells[targetRow, cellRange.FirstColumn + j].Value = rowData[j];
        }
    }
}
