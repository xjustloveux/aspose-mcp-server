using Aspose.Cells;
using AsposeMcpServer.Helpers.Excel;

namespace AsposeMcpServer.Tests.Helpers.Excel;

public class ExcelNamedRangeHelperTests
{
    #region ParseRangeWithSheetReference Tests - Valid Cases

    [Fact]
    public void ParseRangeWithSheetReference_WithValidRange_ReturnsRange()
    {
        var workbook = new Workbook();
        workbook.Worksheets[0].Name = "Sheet1";

        var result = ExcelNamedRangeHelper.ParseRangeWithSheetReference(workbook, "Sheet1!A1:B2");

        Assert.NotNull(result);
    }

    [Fact]
    public void ParseRangeWithSheetReference_WithQuotedSheetName_ReturnsRange()
    {
        var workbook = new Workbook();
        workbook.Worksheets[0].Name = "My Sheet";

        var result = ExcelNamedRangeHelper.ParseRangeWithSheetReference(workbook, "'My Sheet'!A1:C3");

        Assert.NotNull(result);
    }

    [Fact]
    public void ParseRangeWithSheetReference_WithSingleCell_ReturnsRange()
    {
        var workbook = new Workbook();
        workbook.Worksheets[0].Name = "Data";

        var result = ExcelNamedRangeHelper.ParseRangeWithSheetReference(workbook, "Data!A1");

        Assert.NotNull(result);
    }

    #endregion

    #region ParseRangeWithSheetReference Tests - Invalid Cases

    [Fact]
    public void ParseRangeWithSheetReference_WithNoExclamation_ThrowsArgumentException()
    {
        var workbook = new Workbook();

        var ex = Assert.Throws<ArgumentException>(() =>
            ExcelNamedRangeHelper.ParseRangeWithSheetReference(workbook, "A1:B2"));

        Assert.Contains("Invalid range format", ex.Message);
        Assert.Contains("Expected format", ex.Message);
    }

    [Fact]
    public void ParseRangeWithSheetReference_WithExclamationAtStart_ThrowsArgumentException()
    {
        var workbook = new Workbook();

        var ex = Assert.Throws<ArgumentException>(() =>
            ExcelNamedRangeHelper.ParseRangeWithSheetReference(workbook, "!A1:B2"));

        Assert.Contains("Invalid range format", ex.Message);
    }

    [Fact]
    public void ParseRangeWithSheetReference_WithNonExistentSheet_ThrowsArgumentException()
    {
        var workbook = new Workbook();
        workbook.Worksheets[0].Name = "Sheet1";

        var ex = Assert.Throws<ArgumentException>(() =>
            ExcelNamedRangeHelper.ParseRangeWithSheetReference(workbook, "NonExistent!A1:B2"));

        Assert.Contains("Worksheet 'NonExistent' not found", ex.Message);
    }

    #endregion

    #region CreateRangeFromAddress Tests

    [Fact]
    public void CreateRangeFromAddress_WithRangeAddress_ReturnsRange()
    {
        var workbook = new Workbook();
        var cells = workbook.Worksheets[0].Cells;

        var result = ExcelNamedRangeHelper.CreateRangeFromAddress(cells, "A1:C3");

        Assert.NotNull(result);
    }

    [Fact]
    public void CreateRangeFromAddress_WithSingleCellAddress_ReturnsRange()
    {
        var workbook = new Workbook();
        var cells = workbook.Worksheets[0].Cells;

        var result = ExcelNamedRangeHelper.CreateRangeFromAddress(cells, "B5");

        Assert.NotNull(result);
    }

    [Fact]
    public void CreateRangeFromAddress_WithWhitespace_TrimsAndReturnsRange()
    {
        var workbook = new Workbook();
        var cells = workbook.Worksheets[0].Cells;

        var result = ExcelNamedRangeHelper.CreateRangeFromAddress(cells, "  A1 : B2  ");

        Assert.NotNull(result);
    }

    [Fact]
    public void CreateRangeFromAddress_WithColumnRange_ReturnsRange()
    {
        var workbook = new Workbook();
        var cells = workbook.Worksheets[0].Cells;

        var result = ExcelNamedRangeHelper.CreateRangeFromAddress(cells, "A1:A10");

        Assert.NotNull(result);
    }

    [Fact]
    public void CreateRangeFromAddress_WithRowRange_ReturnsRange()
    {
        var workbook = new Workbook();
        var cells = workbook.Worksheets[0].Cells;

        var result = ExcelNamedRangeHelper.CreateRangeFromAddress(cells, "A1:Z1");

        Assert.NotNull(result);
    }

    #endregion
}
