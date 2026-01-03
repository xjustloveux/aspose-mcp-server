using AsposeMcpServer.Tests.Helpers;
using AsposeMcpServer.Tools.Excel;

namespace AsposeMcpServer.Tests.Tools.Excel;

public class ExcelGetCellAddressToolTests : ExcelTestBase
{
    private readonly ExcelGetCellAddressTool _tool = new();

    #region General Tests

    [Fact]
    public void ConvertA1ToIndex_ShouldReturnCorrectIndex()
    {
        var result = _tool.Execute("A1");
        Assert.Equal("A1 = Row 0, Column 0", result);
    }

    [Fact]
    public void ConvertB2ToIndex_ShouldReturnCorrectIndex()
    {
        var result = _tool.Execute("B2");
        Assert.Equal("B2 = Row 1, Column 1", result);
    }

    [Fact]
    public void ConvertAA100ToIndex_ShouldReturnCorrectIndex()
    {
        var result = _tool.Execute("AA100");
        Assert.Equal("AA100 = Row 99, Column 26", result);
    }

    [Fact]
    public void ConvertIndexToA1_ShouldReturnCorrectAddress()
    {
        var result = _tool.Execute(row: 0, column: 0);
        Assert.Equal("A1 = Row 0, Column 0", result);
    }

    [Fact]
    public void ConvertIndexToB2_ShouldReturnCorrectAddress()
    {
        var result = _tool.Execute(row: 1, column: 1);
        Assert.Equal("B2 = Row 1, Column 1", result);
    }

    [Fact]
    public void ConvertLargeIndex_ShouldReturnCorrectAddress()
    {
        var result = _tool.Execute(row: 999, column: 100);
        Assert.Contains("Row 999", result);
        Assert.Contains("Column 100", result);
    }

    [Fact]
    public void MaxValidRow_ShouldSucceed()
    {
        var result = _tool.Execute(row: 1048575, column: 0);
        Assert.Contains("Row 1048575", result);
    }

    [Fact]
    public void MaxValidColumn_ShouldSucceed()
    {
        var result = _tool.Execute(row: 0, column: 16383);
        Assert.Contains("Column 16383", result);
        Assert.Contains("XFD", result);
    }

    [Fact]
    public void LowercaseCellAddress_ShouldWork()
    {
        var result = _tool.Execute("b2");
        Assert.Contains("Row 1", result);
        Assert.Contains("Column 1", result);
    }

    #endregion

    #region Exception Tests

    [Fact]
    public void BothCellAddressAndRowColumn_ShouldThrowException()
    {
        var exception = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("B2", 0, 0));
        Assert.Contains("Cannot specify both", exception.Message);
    }

    [Fact]
    public void NeitherCellAddressNorRowColumn_ShouldThrowException()
    {
        var exception = Assert.Throws<ArgumentException>(() =>
            _tool.Execute());
        Assert.Contains("Must specify either", exception.Message);
    }

    [Fact]
    public void OnlyRow_ShouldThrowException()
    {
        var exception = Assert.Throws<ArgumentException>(() =>
            _tool.Execute(row: 5));
        Assert.Contains("Both row and column must be specified", exception.Message);
    }

    [Fact]
    public void OnlyColumn_ShouldThrowException()
    {
        var exception = Assert.Throws<ArgumentException>(() =>
            _tool.Execute(column: 5));
        Assert.Contains("Both row and column must be specified", exception.Message);
    }

    [Fact]
    public void NegativeRow_ShouldThrowException()
    {
        var exception = Assert.Throws<ArgumentException>(() =>
            _tool.Execute(row: -1, column: 0));
        Assert.Contains("out of range", exception.Message);
    }

    [Fact]
    public void NegativeColumn_ShouldThrowException()
    {
        var exception = Assert.Throws<ArgumentException>(() =>
            _tool.Execute(row: 0, column: -1));
        Assert.Contains("out of range", exception.Message);
    }

    [Fact]
    public void RowExceedsMax_ShouldThrowException()
    {
        var exception = Assert.Throws<ArgumentException>(() =>
            _tool.Execute(row: 1048576, column: 0));
        Assert.Contains("out of range", exception.Message);
    }

    [Fact]
    public void ColumnExceedsMax_ShouldThrowException()
    {
        var exception = Assert.Throws<ArgumentException>(() =>
            _tool.Execute(row: 0, column: 16384));
        Assert.Contains("out of range", exception.Message);
    }

    #endregion

    // Note: This tool does not support session, so no Session ID Tests region
}