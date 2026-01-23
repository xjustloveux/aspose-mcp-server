using Aspose.Cells;
using Cell = Aspose.Cells.Cell;
using Range = Aspose.Cells.Range;

namespace AsposeMcpServer.Helpers.Excel;

/// <summary>
///     Helper class for common Excel operations to reduce code duplication
/// </summary>
public static class ExcelHelper
{
    /// <summary>
    ///     Validates sheet index and throws exception if invalid
    /// </summary>
    /// <param name="sheetIndex">Sheet index to validate</param>
    /// <param name="workbook">Workbook to check against</param>
    /// <exception cref="ArgumentException">Thrown if sheet index is invalid</exception>
    public static void ValidateSheetIndex(int sheetIndex, Workbook workbook)
    {
        if (sheetIndex < 0 || sheetIndex >= workbook.Worksheets.Count)
            throw new ArgumentException(
                $"Sheet index {sheetIndex} is out of range (workbook has {workbook.Worksheets.Count} worksheets)");
    }

    /// <summary>
    ///     Validates sheet index and throws exception if invalid (with custom error message)
    /// </summary>
    /// <param name="sheetIndex">Sheet index to validate</param>
    /// <param name="workbook">Workbook to check against</param>
    /// <param name="customMessage">Custom error message prefix</param>
    /// <exception cref="ArgumentException">Thrown if sheet index is invalid</exception>
    public static void ValidateSheetIndex(int sheetIndex, Workbook workbook, string customMessage)
    {
        if (sheetIndex < 0 || sheetIndex >= workbook.Worksheets.Count)
            throw new ArgumentException(
                $"{customMessage}: Sheet index {sheetIndex} is out of range (workbook has {workbook.Worksheets.Count} worksheets)");
    }

    /// <summary>
    ///     Gets a worksheet with validation
    /// </summary>
    /// <param name="workbook">Workbook to get worksheet from</param>
    /// <param name="sheetIndex">Sheet index</param>
    /// <returns>Worksheet</returns>
    /// <exception cref="ArgumentException">Thrown if sheet index is invalid</exception>
    public static Worksheet GetWorksheet(Workbook workbook, int sheetIndex)
    {
        ValidateSheetIndex(sheetIndex, workbook);
        return workbook.Worksheets[sheetIndex];
    }

    /// <summary>
    ///     Gets a worksheet with validation (with custom error message)
    /// </summary>
    /// <param name="workbook">Workbook to get worksheet from</param>
    /// <param name="sheetIndex">Sheet index</param>
    /// <param name="customMessage">Custom error message prefix</param>
    /// <returns>Worksheet</returns>
    /// <exception cref="ArgumentException">Thrown if sheet index is invalid</exception>
    public static Worksheet GetWorksheet(Workbook workbook, int sheetIndex, string customMessage)
    {
        ValidateSheetIndex(sheetIndex, workbook, customMessage);
        return workbook.Worksheets[sheetIndex];
    }

    /// <summary>
    ///     Creates a range with validation and unified error handling
    ///     This method wraps CreateRange with try-catch to provide consistent error messages
    /// </summary>
    /// <param name="cells">Cells collection to create range from</param>
    /// <param name="range">Range string (e.g., "A1:C5", "Sheet1!A1:C5")</param>
    /// <returns>Range object</returns>
    /// <exception cref="ArgumentException">Thrown if range format is invalid</exception>
    public static Range CreateRange(Cells cells, string range)
    {
        try
        {
            return cells.CreateRange(range);
        }
        catch (Exception ex)
        {
            if (range.Contains(':'))
            {
                var parts = range.Split(':');
                if (parts.Length == 2)
                {
                    var startCell = parts[0].Trim();
                    var endCell = parts[1].Trim();
                    throw new ArgumentException(
                        $"Invalid range format: '{range}'. Range exceeds Excel limits (valid range: A1:XFD1048576). Start cell: '{startCell}', End cell: '{endCell}'. Error: {ex.Message}");
                }
            }

            throw new ArgumentException(
                $"Invalid range format: '{range}'. Range exceeds Excel limits (valid range: A1:XFD1048576). Error: {ex.Message}");
        }
    }

    /// <summary>
    ///     Creates a range with validation and unified error handling (for multiple ranges)
    ///     This method wraps CreateRange with try-catch to provide consistent error messages
    /// </summary>
    /// <param name="cells">Cells collection to create range from</param>
    /// <param name="range">Range string (e.g., "A1:C5", "Sheet1!A1:C5")</param>
    /// <param name="rangeDescription">Description of the range for error message (e.g., "source range", "destination range")</param>
    /// <returns>Range object</returns>
    /// <exception cref="ArgumentException">Thrown if range format is invalid</exception>
    public static Range CreateRange(Cells cells, string range, string rangeDescription)
    {
        try
        {
            return cells.CreateRange(range);
        }
        catch (Exception ex)
        {
            if (range.Contains(':'))
            {
                var parts = range.Split(':');
                if (parts.Length == 2)
                {
                    var startCell = parts[0].Trim();
                    var endCell = parts[1].Trim();
                    throw new ArgumentException(
                        $"Invalid {rangeDescription} format: '{range}'. Range exceeds Excel limits (valid range: A1:XFD1048576). Start cell: '{startCell}', End cell: '{endCell}'. Error: {ex.Message}");
                }
            }

            throw new ArgumentException(
                $"Invalid {rangeDescription} format: '{range}'. Range exceeds Excel limits (valid range: A1:XFD1048576). Error: {ex.Message}");
        }
    }

    /// <summary>
    ///     Sets cell value with automatic type conversion (number, boolean, date, or string).
    ///     This ensures formulas can correctly identify numeric values.
    /// </summary>
    /// <param name="cell">Cell to set value on.</param>
    /// <param name="value">String value to parse and set.</param>
    public static void SetCellValue(Cell cell, string value)
    {
        cell.PutValue(ValueHelper.ParseValue(value));
    }
}
