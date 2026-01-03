using System.ComponentModel;
using Aspose.Cells;
using ModelContextProtocol.Server;

namespace AsposeMcpServer.Tools.Excel;

/// <summary>
///     Tool for converting between Excel cell address formats (A1 notation and row/column index).
/// </summary>
[McpServerToolType]
public class ExcelGetCellAddressTool
{
    /// <summary>
    ///     Maximum number of rows in an Excel worksheet.
    /// </summary>
    private const int MaxExcelRows = 1048576;

    /// <summary>
    ///     Maximum number of columns in an Excel worksheet.
    /// </summary>
    private const int MaxExcelColumns = 16384;

    [McpServerTool(Name = "excel_get_cell_address")]
    [Description(@"Convert between cell address formats (A1 notation and row/column index).

Usage examples:
- Convert A1 to index: excel_get_cell_address(cellAddress='B2') returns row/column index
- Convert index to A1: excel_get_cell_address(row=1, column=1) returns 'B2'
- Validate address: excel_get_cell_address(cellAddress='AA100') validates and returns info")]
    public string Execute(
        [Description("Cell address in A1 notation (e.g., 'A1', 'B2', 'AA100'). Use this OR row/column, not both.")]
        string? cellAddress = null,
        [Description("Row index (0-based, 0 to 1048575). Use with column parameter.")]
        int? row = null,
        [Description("Column index (0-based, 0 to 16383). Use with row parameter.")]
        int? column = null)
    {
        var hasCellAddress = !string.IsNullOrWhiteSpace(cellAddress);
        var hasRowColumn = row.HasValue && column.HasValue;
        var hasPartialRowColumn = row.HasValue != column.HasValue;

        if (hasPartialRowColumn)
            throw new ArgumentException("Both row and column must be specified together.");

        if (hasCellAddress && hasRowColumn)
            throw new ArgumentException(
                "Cannot specify both cellAddress and row/column. Use one or the other.");

        if (!hasCellAddress && !hasRowColumn)
            throw new ArgumentException(
                "Must specify either cellAddress or both row and column parameters.");

        int finalRow, finalCol;

        if (hasRowColumn)
        {
            finalRow = row!.Value;
            finalCol = column!.Value;
        }
        else
        {
            CellsHelper.CellNameToIndex(cellAddress!, out finalRow, out finalCol);
        }

        ValidateIndexBounds(finalRow, finalCol);

        var a1Notation = CellsHelper.CellIndexToName(finalRow, finalCol);
        return $"{a1Notation} = Row {finalRow}, Column {finalCol}";
    }

    /// <summary>
    ///     Validates that row and column indices are within Excel's valid range.
    /// </summary>
    /// <param name="row">The row index to validate.</param>
    /// <param name="column">The column index to validate.</param>
    /// <exception cref="ArgumentException">Thrown when the index is out of valid range.</exception>
    private static void ValidateIndexBounds(int row, int column)
    {
        if (row < 0 || row >= MaxExcelRows)
            throw new ArgumentException(
                $"Row index {row} is out of range. Valid range: 0 to {MaxExcelRows - 1}.");

        if (column < 0 || column >= MaxExcelColumns)
            throw new ArgumentException(
                $"Column index {column} is out of range. Valid range: 0 to {MaxExcelColumns - 1}.");
    }
}