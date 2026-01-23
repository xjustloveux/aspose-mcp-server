using System.Text.RegularExpressions;
using Aspose.Cells;

namespace AsposeMcpServer.Helpers.Excel;

/// <summary>
///     Helper class for Excel cell operations in handlers.
/// </summary>
public static partial class ExcelCellHelper
{
    /// <summary>
    ///     Regex pattern for validating Excel cell addresses (e.g., A1, B2, AA100).
    /// </summary>
    private static readonly Regex CellAddressRegex = CellAddressRegexGenerated();

    /// <summary>
    ///     Validates the cell address format.
    /// </summary>
    /// <param name="cell">The cell address to validate.</param>
    /// <exception cref="ArgumentException">Thrown when the cell address format is invalid.</exception>
    public static void ValidateCellAddress(string cell)
    {
        if (string.IsNullOrEmpty(cell))
            throw new ArgumentException("cell is required");

        if (!CellAddressRegex.IsMatch(cell))
            throw new ArgumentException(
                $"Invalid cell address format: '{cell}'. Expected format like 'A1', 'B2', 'AA100'");
    }

    /// <summary>
    ///     Gets a cell from a worksheet with validation.
    /// </summary>
    /// <param name="workbook">The workbook.</param>
    /// <param name="sheetIndex">The sheet index.</param>
    /// <param name="cell">The cell address.</param>
    /// <returns>The cell object.</returns>
    public static Cell GetCell(Workbook workbook, int sheetIndex, string cell)
    {
        var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);
        return worksheet.Cells[cell];
    }

    [GeneratedRegex(@"^[A-Za-z]{1,3}\d+$", RegexOptions.Compiled)]
    private static partial Regex CellAddressRegexGenerated();
}
