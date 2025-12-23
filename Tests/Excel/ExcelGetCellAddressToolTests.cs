using System.Text.Json.Nodes;
using AsposeMcpServer.Tests.Helpers;
using AsposeMcpServer.Tools.Excel;

namespace AsposeMcpServer.Tests.Excel;

public class ExcelGetCellAddressToolTests : ExcelTestBase
{
    private readonly ExcelGetCellAddressTool _tool = new();

    [Fact]
    public async Task ConvertA1ToIndex_ShouldConvertAddress()
    {
        // Arrange
        var arguments = new JsonObject
        {
            ["cellAddress"] = "A1",
            ["convertToIndex"] = true
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.NotNull(result);
        Assert.NotEmpty(result);
        Assert.Contains("0", result); // Row 0
        Assert.Contains("0", result); // Column 0
    }

    [Fact]
    public async Task ConvertIndexToA1_ShouldConvertAddress()
    {
        // Arrange
        var arguments = new JsonObject
        {
            ["row"] = 0,
            ["column"] = 0
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.NotNull(result);
        Assert.NotEmpty(result);
        Assert.Contains("A1", result, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public async Task ConvertB2_ShouldConvertAddress()
    {
        // Arrange
        var arguments = new JsonObject
        {
            ["cellAddress"] = "B2"
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.NotNull(result);
        Assert.NotEmpty(result);
    }
}