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

            List<List<object?>> rows = [];
            var startRow = hasHeader ? cellRange.FirstRow + 1 : cellRange.FirstRow;

            if (hasHeader)
            {
                List<object?> headerRow = [];
                for (var col = cellRange.FirstColumn; col < cellRange.FirstColumn + cellRange.ColumnCount; col++)
                    headerRow.Add(cells[cellRange.FirstRow, col].Value);
                rows.Add(headerRow);
            }

            for (var row = startRow; row < cellRange.FirstRow + cellRange.RowCount; row++)
            {
                List<object?> rowData = [];
                for (var col = cellRange.FirstColumn; col < cellRange.FirstColumn + cellRange.ColumnCount; col++)
                    rowData.Add(cells[row, col].Value);
                rows.Add(rowData);
            }

            var dataRows = hasHeader ? rows.Skip(1).ToList() : rows;
            dataRows.Sort((a, b) =>
            {
                var aVal = a[sortColumn];
                var bVal = b[sortColumn];

                if (aVal == null && bVal == null) return 0;
                if (aVal == null) return ascending ? -1 : 1;
                if (bVal == null) return ascending ? 1 : -1;

                var comparison = Comparer<object>.Default.Compare(aVal, bVal);
                return ascending ? comparison : -comparison;
            });

            if (hasHeader)
            {
                rows = [rows[0]];
                rows.AddRange(dataRows);
            }
            else
            {
                rows = dataRows;
            }

            for (var i = 0; i < rows.Count; i++)
            {
                var rowData = rows[i];
                var targetRow = cellRange.FirstRow + i;
                for (var j = 0; j < rowData.Count; j++)
                    cells[targetRow, cellRange.FirstColumn + j].Value = rowData[j];
            }

            MarkModified(context);

            return Success(
                $"Sorted range {range} by column {sortColumn} ({(ascending ? "ascending" : "descending")}).");
        }
        catch (CellsException ex)
        {
            throw new ArgumentException($"Excel operation failed for range '{range}': {ex.Message}");
        }
    }
}
