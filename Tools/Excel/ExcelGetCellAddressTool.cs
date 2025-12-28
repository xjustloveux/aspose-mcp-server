using System.Text.Json.Nodes;
using Aspose.Cells;
using AsposeMcpServer.Core;

namespace AsposeMcpServer.Tools.Excel;

/// <summary>
///     Tool for converting between Excel cell address formats (A1 notation and row/column index).
/// </summary>
public class ExcelGetCellAddressTool : IAsposeTool
{
    private const int MaxExcelRows = 1048576;
    private const int MaxExcelColumns = 16384;

    /// <summary>
    ///     Gets the description of the tool and its usage examples.
    /// </summary>
    public string Description => @"Convert between cell address formats (A1 notation and row/column index).

Usage examples:
- Convert A1 to index: excel_get_cell_address(cellAddress='B2') returns row/column index
- Convert index to A1: excel_get_cell_address(row=1, column=1) returns 'B2'
- Validate address: excel_get_cell_address(cellAddress='AA100') validates and returns info";

    /// <summary>
    ///     Gets the JSON schema defining the input parameters for the tool.
    /// </summary>
    public object InputSchema => new
    {
        type = "object",
        properties = new
        {
            cellAddress = new
            {
                type = "string",
                description =
                    "Cell address in A1 notation (e.g., 'A1', 'B2', 'AA100'). Use this OR row/column, not both."
            },
            row = new
            {
                type = "number",
                description = "Row index (0-based, 0 to 1048575). Use with column parameter."
            },
            column = new
            {
                type = "number",
                description = "Column index (0-based, 0 to 16383). Use with row parameter."
            }
        },
        required = new string[] { }
    };

    /// <summary>
    ///     Executes the tool operation with the provided JSON arguments.
    /// </summary>
    /// <param name="arguments">JSON arguments object containing operation parameters.</param>
    /// <returns>Result message as a string.</returns>
    public Task<string> ExecuteAsync(JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var cellAddress = ArgumentHelper.GetStringNullable(arguments, "cellAddress");
            var row = ArgumentHelper.GetIntNullable(arguments, "row");
            var column = ArgumentHelper.GetIntNullable(arguments, "column");

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
        });
    }

    /// <summary>
    ///     Validates that row and column indices are within Excel limits.
    /// </summary>
    /// <param name="row">Row index to validate.</param>
    /// <param name="column">Column index to validate.</param>
    /// <exception cref="ArgumentException">Thrown if indices are out of bounds.</exception>
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