using AsposeMcpServer.Results.Excel.CellAddress;
using AsposeMcpServer.Tests.Infrastructure;
using AsposeMcpServer.Tools.Excel;

namespace AsposeMcpServer.Tests.Tools.Excel;

/// <summary>
///     Integration tests for ExcelGetCellAddressTool.
///     Focuses on file I/O and operation routing.
///     Detailed parameter validation and business logic tests are in Handler tests.
/// </summary>
public class ExcelGetCellAddressToolTests : ExcelTestBase
{
    private readonly ExcelGetCellAddressTool _tool = new();

    #region File I/O Smoke Tests

    [Fact]
    public void FromA1_ShouldReturnCorrectIndex()
    {
        var result = _tool.Execute("from_a1", "B2");
        var data = GetResultData<CellAddressResult>(result);
        Assert.Equal("B2", data.A1Notation);
        Assert.Equal(1, data.Row);
        Assert.Equal(1, data.Column);
    }

    [Fact]
    public void FromIndex_ShouldReturnCorrectAddress()
    {
        var result = _tool.Execute("from_index", row: 0, column: 0);
        var data = GetResultData<CellAddressResult>(result);
        Assert.Equal("A1", data.A1Notation);
        Assert.Equal(0, data.Row);
        Assert.Equal(0, data.Column);
    }

    #endregion

    #region Operation Routing

    [Theory]
    [InlineData("FROM_A1")]
    [InlineData("From_A1")]
    [InlineData("from_a1")]
    public void Operation_ShouldBeCaseInsensitive(string operation)
    {
        var result = _tool.Execute(operation, "A1");
        var data = GetResultData<CellAddressResult>(result);
        Assert.Equal("A1", data.A1Notation);
        Assert.Equal(0, data.Row);
        Assert.Equal(0, data.Column);
    }

    [Fact]
    public void Execute_WithInvalidOperation_ShouldThrowArgumentException()
    {
        var ex = Assert.Throws<ArgumentException>(() => _tool.Execute("invalid_operation"));
        Assert.Contains("operation", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    #endregion
}
