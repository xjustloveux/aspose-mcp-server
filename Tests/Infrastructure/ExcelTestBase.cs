using Aspose.Cells;

namespace AsposeMcpServer.Tests.Infrastructure;

/// <summary>
///     Base class for Excel tool tests providing Excel-specific functionality
/// </summary>
public abstract class ExcelTestBase : TestBase
{
    /// <summary>
    ///     Creates a new Excel workbook for testing
    /// </summary>
    protected string CreateExcelWorkbook(string fileName)
    {
        var filePath = CreateTestFilePath(fileName);
        var workbook = new Workbook();
        workbook.Save(filePath);
        return filePath;
    }

    /// <summary>
    ///     Creates an Excel workbook with sample data
    /// </summary>
    protected string CreateExcelWorkbookWithData(string fileName, int rowCount = 5, int columnCount = 3)
    {
        var filePath = CreateTestFilePath(fileName);
        var workbook = new Workbook();
        var worksheet = workbook.Worksheets[0];

        for (var row = 0; row < rowCount; row++)
        for (var col = 0; col < columnCount; col++)
            worksheet.Cells[row, col].Value = $"R{row + 1}C{col + 1}";

        workbook.Save(filePath);
        return filePath;
    }

    /// <summary>
    ///     Verifies that a cell has the expected value
    /// </summary>
    protected static void AssertCellValue(Workbook workbook, int sheetIndex, int row, int column, object expectedValue)
    {
        var worksheet = workbook.Worksheets[sheetIndex];
        var actualValue = worksheet.Cells[row, column].Value;
        Assert.Equal(expectedValue, actualValue);
    }

    /// <summary>
    ///     Checks if Aspose.Cells is running in evaluation mode.
    /// </summary>
    protected new static bool IsEvaluationMode(AsposeLibraryType libraryType = AsposeLibraryType.Cells)
    {
        return TestBase.IsEvaluationMode(libraryType);
    }
}
