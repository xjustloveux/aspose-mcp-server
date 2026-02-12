using Aspose.Cells;
using AsposeMcpServer.Helpers.Excel;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Helpers;

/// <summary>
///     Unit tests for ExcelHelper class
/// </summary>
public class ExcelHelperTests : TestBase
{
    #region ValidateSheetIndex Tests

    [Fact]
    public void ValidateSheetIndex_WithValidIndex_ShouldNotThrow()
    {
        using var workbook = new Workbook();
        workbook.Worksheets.Add("Sheet2");

        var exception = Record.Exception(() => ExcelHelper.ValidateSheetIndex(0, workbook));

        Assert.Null(exception);
    }

    [Fact]
    public void ValidateSheetIndex_WithNegativeIndex_ShouldThrow()
    {
        using var workbook = new Workbook();

        var ex = Assert.Throws<ArgumentException>(() =>
            ExcelHelper.ValidateSheetIndex(-1, workbook));

        Assert.Contains("out of range", ex.Message);
    }

    [Fact]
    public void ValidateSheetIndex_WithIndexExceedingCount_ShouldThrow()
    {
        using var workbook = new Workbook();

        var ex = Assert.Throws<ArgumentException>(() =>
            ExcelHelper.ValidateSheetIndex(5, workbook));

        Assert.Contains("out of range", ex.Message);
        Assert.Contains("1 worksheets", ex.Message);
    }

    [Fact]
    public void ValidateSheetIndex_WithCustomMessage_ShouldIncludeMessage()
    {
        using var workbook = new Workbook();

        var ex = Assert.Throws<ArgumentException>(() =>
            ExcelHelper.ValidateSheetIndex(5, workbook, "Custom error"));

        Assert.Contains("Custom error", ex.Message);
        Assert.Contains("out of range", ex.Message);
    }

    #endregion

    #region GetWorksheet Tests

    [Fact]
    public void GetWorksheet_WithValidIndex_ShouldReturnWorksheet()
    {
        using var workbook = new Workbook();

        var worksheet = ExcelHelper.GetWorksheet(workbook, 0);

        Assert.NotNull(worksheet);
    }

    [Fact]
    public void GetWorksheet_WithInvalidIndex_ShouldThrow()
    {
        using var workbook = new Workbook();

        Assert.Throws<ArgumentException>(() =>
            ExcelHelper.GetWorksheet(workbook, 5));
    }

    [Fact]
    public void GetWorksheet_WithCustomMessage_ShouldThrowWithMessage()
    {
        using var workbook = new Workbook();

        var ex = Assert.Throws<ArgumentException>(() =>
            ExcelHelper.GetWorksheet(workbook, 5, "Source sheet"));

        Assert.Contains("Source sheet", ex.Message);
    }

    #endregion

    #region CreateRange Tests

    [Fact]
    public void CreateRange_WithValidRange_ShouldReturnRange()
    {
        using var workbook = new Workbook();
        var cells = workbook.Worksheets[0].Cells;

        var range = ExcelHelper.CreateRange(cells, "A1:C5");

        Assert.NotNull(range);
    }

    [Fact]
    public void CreateRange_WithSingleCell_ShouldReturnRange()
    {
        using var workbook = new Workbook();
        var cells = workbook.Worksheets[0].Cells;

        var range = ExcelHelper.CreateRange(cells, "A1");

        Assert.NotNull(range);
    }

    [Fact]
    public void CreateRange_WithInvalidRange_ShouldThrowWithDetails()
    {
        using var workbook = new Workbook();
        var cells = workbook.Worksheets[0].Cells;

        var ex = Assert.Throws<ArgumentException>(() =>
            ExcelHelper.CreateRange(cells, "A1:ZZZZZZZ9999999"));

        Assert.Contains("Invalid range format", ex.Message);
        Assert.Contains("Excel limits", ex.Message);
    }

    [Fact]
    public void CreateRange_WithRangeDescription_ShouldIncludeInError()
    {
        using var workbook = new Workbook();
        var cells = workbook.Worksheets[0].Cells;

        var ex = Assert.Throws<ArgumentException>(() =>
            ExcelHelper.CreateRange(cells, "A1:ZZZZZZZ9999999", "source range"));

        Assert.Contains("source range", ex.Message);
    }

    #endregion

    #region SetCellValue Tests

    [Fact]
    public void SetCellValue_WithIntegerString_ShouldSetAsNumber()
    {
        using var workbook = new Workbook();
        var cell = workbook.Worksheets[0].Cells["A1"];

        ExcelHelper.SetCellValue(cell, "123");

        Assert.Equal(123, cell.IntValue);
    }

    [Fact]
    public void SetCellValue_WithDecimalString_ShouldSetAsDouble()
    {
        using var workbook = new Workbook();
        var cell = workbook.Worksheets[0].Cells["A1"];

        ExcelHelper.SetCellValue(cell, "123.45");

        Assert.Equal(123.45, cell.DoubleValue);
    }

    [Fact]
    public void SetCellValue_WithBooleanString_ShouldSetAsBoolean()
    {
        using var workbook = new Workbook();
        var cell = workbook.Worksheets[0].Cells["A1"];

        ExcelHelper.SetCellValue(cell, "true");

        Assert.True(cell.BoolValue);
    }

    [Fact]
    public void SetCellValue_WithTextString_ShouldSetAsString()
    {
        using var workbook = new Workbook();
        var cell = workbook.Worksheets[0].Cells["A1"];

        ExcelHelper.SetCellValue(cell, "Hello World");

        Assert.Equal("Hello World", cell.StringValue);
    }

    #endregion
}
