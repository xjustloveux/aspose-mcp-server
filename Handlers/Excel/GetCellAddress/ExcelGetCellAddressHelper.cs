namespace AsposeMcpServer.Handlers.Excel.GetCellAddress;

/// <summary>
///     Helper class for Excel cell address operations.
/// </summary>
public static class ExcelGetCellAddressHelper
{
    /// <summary>
    ///     Maximum number of rows in an Excel worksheet.
    /// </summary>
    public const int MaxExcelRows = 1048576;

    /// <summary>
    ///     Maximum number of columns in an Excel worksheet.
    /// </summary>
    public const int MaxExcelColumns = 16384;

    /// <summary>
    ///     Validates that row and column indices are within Excel's valid range.
    /// </summary>
    /// <param name="row">The row index to validate.</param>
    /// <param name="column">The column index to validate.</param>
    /// <exception cref="ArgumentException">Thrown when the index is out of valid range.</exception>
    public static void ValidateIndexBounds(int row, int column)
    {
        if (row < 0 || row >= MaxExcelRows)
            throw new ArgumentException(
                $"Row index {row} is out of range. Valid range: 0 to {MaxExcelRows - 1}.");

        if (column < 0 || column >= MaxExcelColumns)
            throw new ArgumentException(
                $"Column index {column} is out of range. Valid range: 0 to {MaxExcelColumns - 1}.");
    }
}
