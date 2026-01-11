using AsposeMcpServer.Tests.Helpers;
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
        Assert.Contains("Row 1", result);
        Assert.Contains("Column 1", result);
    }

    [Fact]
    public void FromIndex_ShouldReturnCorrectAddress()
    {
        var result = _tool.Execute("from_index", row: 0, column: 0);
        Assert.Contains("A1", result);
        Assert.Contains("Row 0", result);
        Assert.Contains("Column 0", result);
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
        Assert.Contains("Row", result);
        Assert.Contains("Column", result);
    }

    [Fact]
    public void Execute_WithInvalidOperation_ShouldThrowArgumentException()
    {
        var ex = Assert.Throws<ArgumentException>(() => _tool.Execute("invalid_operation"));
        Assert.Contains("operation", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    #endregion
}
