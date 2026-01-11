using Aspose.Cells;

namespace AsposeMcpServer.Handlers.Excel.NamedRange;

/// <summary>
///     Helper class for Excel named range operations.
/// </summary>
public static class ExcelNamedRangeHelper
{
    /// <summary>
    ///     Parses a range address that includes a sheet reference (e.g., "Sheet1!A1:B2").
    /// </summary>
    /// <param name="workbook">The workbook containing the worksheets.</param>
    /// <param name="rangeAddress">The range address with sheet reference (e.g., 'Sheet1!A1:B2').</param>
    /// <returns>The Range object corresponding to the address.</returns>
    /// <exception cref="ArgumentException">Thrown when the range format is invalid or the worksheet is not found.</exception>
    public static Aspose.Cells.Range ParseRangeWithSheetReference(Workbook workbook, string rangeAddress)
    {
        var exclamationIndex = rangeAddress.LastIndexOf('!');
        if (exclamationIndex <= 0)
            throw new ArgumentException($"Invalid range format: '{rangeAddress}'. Expected format: 'SheetName!A1:C1'");

        var sheetRef = rangeAddress[..exclamationIndex].Trim().Trim('\'');
        var cellRange = rangeAddress[(exclamationIndex + 1)..].Trim();

        Worksheet? targetSheet = null;
        foreach (var ws in workbook.Worksheets)
            if (ws.Name == sheetRef)
            {
                targetSheet = ws;
                break;
            }

        if (targetSheet == null)
            throw new ArgumentException($"Worksheet '{sheetRef}' not found.");

        return CreateRangeFromAddress(targetSheet.Cells, cellRange);
    }

    /// <summary>
    ///     Creates a Range object from a cell address (e.g., "A1:B2" or "A1").
    /// </summary>
    /// <param name="cells">The Cells collection of the worksheet.</param>
    /// <param name="address">The cell address (e.g., 'A1:B2' or 'A1').</param>
    /// <returns>The Range object corresponding to the address.</returns>
    public static Aspose.Cells.Range CreateRangeFromAddress(Cells cells, string address)
    {
        var colonIndex = address.IndexOf(':');
        if (colonIndex > 0)
        {
            var startCell = address[..colonIndex].Trim();
            var endCell = address[(colonIndex + 1)..].Trim();
            return cells.CreateRange(startCell, endCell);
        }

        return cells.CreateRange(address.Trim(), address.Trim());
    }
}
