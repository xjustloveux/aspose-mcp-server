using AsposeMcpServer.Tests.Helpers;
using AsposeMcpServer.Tools.Excel;

namespace AsposeMcpServer.Tests.Tools.Excel;

public class ExcelGetCellAddressToolTests : ExcelTestBase
{
    private readonly ExcelGetCellAddressTool _tool = new();

    #region General

    [Theory]
    [InlineData("A1", 0, 0)]
    [InlineData("B2", 1, 1)]
    [InlineData("C10", 9, 2)]
    [InlineData("Z1", 0, 25)]
    [InlineData("AA1", 0, 26)]
    [InlineData("AA100", 99, 26)]
    [InlineData("AZ50", 49, 51)]
    public void ConvertCellAddressToIndex_ShouldReturnCorrectIndex(string cellAddress, int expectedRow, int expectedCol)
    {
        var result = _tool.Execute(cellAddress);
        Assert.Contains($"Row {expectedRow}", result);
        Assert.Contains($"Column {expectedCol}", result);
    }

    [Theory]
    [InlineData(0, 0, "A1")]
    [InlineData(1, 1, "B2")]
    [InlineData(9, 2, "C10")]
    [InlineData(0, 25, "Z1")]
    [InlineData(0, 26, "AA1")]
    [InlineData(99, 26, "AA100")]
    [InlineData(999, 100, "CW1000")]
    public void ConvertIndexToCellAddress_ShouldReturnCorrectAddress(int row, int column, string expectedAddress)
    {
        var result = _tool.Execute(row: row, column: column);
        Assert.Contains(expectedAddress, result);
        Assert.Contains($"Row {row}", result);
        Assert.Contains($"Column {column}", result);
    }

    [Theory]
    [InlineData("a1")]
    [InlineData("A1")]
    [InlineData("b2")]
    [InlineData("B2")]
    [InlineData("aa100")]
    [InlineData("AA100")]
    public void ConvertCellAddress_ShouldBeCaseInsensitive(string cellAddress)
    {
        var result = _tool.Execute(cellAddress);
        Assert.Contains("Row", result);
        Assert.Contains("Column", result);
    }

    [Fact]
    public void MaxValidRow_ShouldSucceed()
    {
        var result = _tool.Execute(row: 1048575, column: 0);
        Assert.Contains("Row 1048575", result);
        Assert.Contains("A1048576", result);
    }

    [Fact]
    public void MaxValidColumn_ShouldSucceed()
    {
        var result = _tool.Execute(row: 0, column: 16383);
        Assert.Contains("Column 16383", result);
        Assert.Contains("XFD", result);
    }

    [Fact]
    public void MaxValidCell_ShouldSucceed()
    {
        var result = _tool.Execute(row: 1048575, column: 16383);
        Assert.Contains("Row 1048575", result);
        Assert.Contains("Column 16383", result);
        Assert.Contains("XFD1048576", result);
    }

    #endregion

    #region Exception

    [Fact]
    public void Execute_WithBothCellAddressAndRowColumn_ShouldThrowArgumentException()
    {
        var ex = Assert.Throws<ArgumentException>(() => _tool.Execute("B2", 0, 0));
        Assert.Contains("Cannot specify both", ex.Message);
    }

    [Fact]
    public void Execute_WithNeitherCellAddressNorRowColumn_ShouldThrowArgumentException()
    {
        var ex = Assert.Throws<ArgumentException>(() => _tool.Execute());
        Assert.Contains("Must specify either", ex.Message);
    }

    [Fact]
    public void Execute_WithOnlyRow_ShouldThrowArgumentException()
    {
        var ex = Assert.Throws<ArgumentException>(() => _tool.Execute(row: 5));
        Assert.Contains("Both row and column must be specified", ex.Message);
    }

    [Fact]
    public void Execute_WithOnlyColumn_ShouldThrowArgumentException()
    {
        var ex = Assert.Throws<ArgumentException>(() => _tool.Execute(column: 5));
        Assert.Contains("Both row and column must be specified", ex.Message);
    }

    [Fact]
    public void Execute_WithNegativeRow_ShouldThrowArgumentException()
    {
        var ex = Assert.Throws<ArgumentException>(() => _tool.Execute(row: -1, column: 0));
        Assert.Contains("out of range", ex.Message);
    }

    [Fact]
    public void Execute_WithNegativeColumn_ShouldThrowArgumentException()
    {
        var ex = Assert.Throws<ArgumentException>(() => _tool.Execute(row: 0, column: -1));
        Assert.Contains("out of range", ex.Message);
    }

    [Fact]
    public void Execute_WithRowExceedsMax_ShouldThrowArgumentException()
    {
        var ex = Assert.Throws<ArgumentException>(() => _tool.Execute(row: 1048576, column: 0));
        Assert.Contains("out of range", ex.Message);
    }

    [Fact]
    public void Execute_WithColumnExceedsMax_ShouldThrowArgumentException()
    {
        var ex = Assert.Throws<ArgumentException>(() => _tool.Execute(row: 0, column: 16384));
        Assert.Contains("out of range", ex.Message);
    }

    #endregion
}