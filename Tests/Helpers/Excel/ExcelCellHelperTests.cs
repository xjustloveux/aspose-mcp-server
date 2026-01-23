using Aspose.Cells;
using AsposeMcpServer.Helpers.Excel;

namespace AsposeMcpServer.Tests.Helpers.Excel;

public class ExcelCellHelperTests
{
    #region ValidateCellAddress Tests

    [Theory]
    [InlineData("A1")]
    [InlineData("B2")]
    [InlineData("Z99")]
    [InlineData("AA1")]
    [InlineData("AA100")]
    [InlineData("ABC123")]
    [InlineData("a1")]
    [InlineData("ab99")]
    public void ValidateCellAddress_WithValidAddress_DoesNotThrow(string address)
    {
        var ex = Record.Exception(() => ExcelCellHelper.ValidateCellAddress(address));
        Assert.Null(ex);
    }

    [Theory]
    [InlineData("")]
    [InlineData(null)]
    public void ValidateCellAddress_WithEmptyOrNull_ThrowsArgumentException(string? address)
    {
        var ex = Assert.Throws<ArgumentException>(() => ExcelCellHelper.ValidateCellAddress(address!));
        Assert.Contains("cell is required", ex.Message);
    }

    [Theory]
    [InlineData("1A")]
    [InlineData("A")]
    [InlineData("1")]
    [InlineData("A1A")]
    [InlineData("A-1")]
    [InlineData("A.1")]
    [InlineData("A:1")]
    [InlineData("A 1")]
    [InlineData("$A$1")]
    [InlineData("ABCD1")]
    public void ValidateCellAddress_WithInvalidFormat_ThrowsArgumentException(string address)
    {
        var ex = Assert.Throws<ArgumentException>(() => ExcelCellHelper.ValidateCellAddress(address));
        Assert.Contains("Invalid cell address format", ex.Message);
    }

    #endregion

    #region GetCell Tests

    [Fact]
    public void GetCell_WithValidParameters_ReturnsCell()
    {
        var workbook = new Workbook();
        workbook.Worksheets[0].Cells["A1"].PutValue("Test");

        var cell = ExcelCellHelper.GetCell(workbook, 0, "A1");

        Assert.NotNull(cell);
        Assert.Equal("Test", cell.StringValue);
    }

    [Fact]
    public void GetCell_WithValidSheetIndex_ReturnsCorrectCell()
    {
        var workbook = new Workbook();
        workbook.Worksheets.Add("Sheet2");
        workbook.Worksheets[1].Cells["B2"].PutValue("Sheet2Value");

        var cell = ExcelCellHelper.GetCell(workbook, 1, "B2");

        Assert.NotNull(cell);
        Assert.Equal("Sheet2Value", cell.StringValue);
    }

    [Fact]
    public void GetCell_WithInvalidSheetIndex_ThrowsException()
    {
        var workbook = new Workbook();

        Assert.ThrowsAny<Exception>(() => ExcelCellHelper.GetCell(workbook, 99, "A1"));
    }

    [Theory]
    [InlineData("A1")]
    [InlineData("Z100")]
    [InlineData("AA50")]
    public void GetCell_WithVariousAddresses_ReturnsCell(string address)
    {
        var workbook = new Workbook();

        var cell = ExcelCellHelper.GetCell(workbook, 0, address);

        Assert.NotNull(cell);
    }

    #endregion
}
