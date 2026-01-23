using Aspose.Cells;

namespace AsposeMcpServer.Tests.Infrastructure;

/// <summary>
///     Base class for Excel Handler tests providing Excel-specific test infrastructure.
/// </summary>
public abstract class ExcelHandlerTestBase : HandlerTestBase<Workbook>
{
    /// <summary>
    ///     Creates a new empty Excel workbook for testing.
    /// </summary>
    /// <returns>A new empty Workbook instance.</returns>
    protected static Workbook CreateEmptyWorkbook()
    {
        return new Workbook();
    }

    /// <summary>
    ///     Creates an Excel workbook with sample data.
    /// </summary>
    /// <param name="data">2D array of cell values.</param>
    /// <returns>A Workbook with the specified data.</returns>
    protected static Workbook CreateWorkbookWithData(object[,] data)
    {
        var workbook = new Workbook();
        var sheet = workbook.Worksheets[0];

        for (var row = 0; row < data.GetLength(0); row++)
        for (var col = 0; col < data.GetLength(1); col++)
            sheet.Cells[row, col].Value = data[row, col];

        return workbook;
    }

    /// <summary>
    ///     Gets a cell value from the first worksheet.
    /// </summary>
    /// <param name="workbook">The workbook.</param>
    /// <param name="row">The row index (0-based).</param>
    /// <param name="col">The column index (0-based).</param>
    /// <returns>The cell value.</returns>
    protected static object? GetCellValue(Workbook workbook, int row, int col)
    {
        return workbook.Worksheets[0].Cells[row, col].Value;
    }

    /// <summary>
    ///     Asserts that a cell contains the expected value.
    /// </summary>
    /// <param name="workbook">The workbook.</param>
    /// <param name="row">The row index (0-based).</param>
    /// <param name="col">The column index (0-based).</param>
    /// <param name="expectedValue">The expected value.</param>
    protected static void AssertCellValue(Workbook workbook, int row, int col, object? expectedValue)
    {
        var actualValue = GetCellValue(workbook, row, col);
        Assert.Equal(expectedValue, actualValue);
    }

    /// <summary>
    ///     Creates an Excel workbook with multiple worksheets.
    /// </summary>
    /// <param name="count">The number of worksheets to create.</param>
    /// <returns>A Workbook with the specified number of worksheets.</returns>
    protected static Workbook CreateWorkbookWithSheets(int count)
    {
        var workbook = new Workbook();
        for (var i = 1; i < count; i++) workbook.Worksheets.Add($"Sheet{i + 1}");
        return workbook;
    }
}
