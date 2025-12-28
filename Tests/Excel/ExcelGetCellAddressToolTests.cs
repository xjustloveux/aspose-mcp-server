using System.Text.Json.Nodes;
using AsposeMcpServer.Tests.Helpers;
using AsposeMcpServer.Tools.Excel;

namespace AsposeMcpServer.Tests.Excel;

public class ExcelGetCellAddressToolTests : ExcelTestBase
{
    private readonly ExcelGetCellAddressTool _tool = new();

    [Fact]
    public async Task ConvertA1ToIndex_ShouldReturnCorrectIndex()
    {
        // Arrange
        var arguments = new JsonObject
        {
            ["cellAddress"] = "A1"
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.Equal("A1 = Row 0, Column 0", result);
    }

    [Fact]
    public async Task ConvertB2ToIndex_ShouldReturnCorrectIndex()
    {
        // Arrange
        var arguments = new JsonObject
        {
            ["cellAddress"] = "B2"
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.Equal("B2 = Row 1, Column 1", result);
    }

    [Fact]
    public async Task ConvertAA100ToIndex_ShouldReturnCorrectIndex()
    {
        // Arrange
        var arguments = new JsonObject
        {
            ["cellAddress"] = "AA100"
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.Equal("AA100 = Row 99, Column 26", result);
    }

    [Fact]
    public async Task ConvertIndexToA1_ShouldReturnCorrectAddress()
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
        Assert.Equal("A1 = Row 0, Column 0", result);
    }

    [Fact]
    public async Task ConvertIndexToB2_ShouldReturnCorrectAddress()
    {
        // Arrange
        var arguments = new JsonObject
        {
            ["row"] = 1,
            ["column"] = 1
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.Equal("B2 = Row 1, Column 1", result);
    }

    [Fact]
    public async Task ConvertLargeIndex_ShouldReturnCorrectAddress()
    {
        // Arrange
        var arguments = new JsonObject
        {
            ["row"] = 999,
            ["column"] = 100
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.Contains("Row 999", result);
        Assert.Contains("Column 100", result);
    }

    [Fact]
    public async Task BothCellAddressAndRowColumn_ShouldThrowException()
    {
        // Arrange
        var arguments = new JsonObject
        {
            ["cellAddress"] = "B2",
            ["row"] = 0,
            ["column"] = 0
        };

        // Act & Assert
        var exception = await Assert.ThrowsAsync<ArgumentException>(() => _tool.ExecuteAsync(arguments));
        Assert.Contains("Cannot specify both", exception.Message);
    }

    [Fact]
    public async Task NeitherCellAddressNorRowColumn_ShouldThrowException()
    {
        // Arrange
        var arguments = new JsonObject();

        // Act & Assert
        var exception = await Assert.ThrowsAsync<ArgumentException>(() => _tool.ExecuteAsync(arguments));
        Assert.Contains("Must specify either", exception.Message);
    }

    [Fact]
    public async Task OnlyRow_ShouldThrowException()
    {
        // Arrange
        var arguments = new JsonObject
        {
            ["row"] = 5
        };

        // Act & Assert
        var exception = await Assert.ThrowsAsync<ArgumentException>(() => _tool.ExecuteAsync(arguments));
        Assert.Contains("Both row and column must be specified", exception.Message);
    }

    [Fact]
    public async Task OnlyColumn_ShouldThrowException()
    {
        // Arrange
        var arguments = new JsonObject
        {
            ["column"] = 5
        };

        // Act & Assert
        var exception = await Assert.ThrowsAsync<ArgumentException>(() => _tool.ExecuteAsync(arguments));
        Assert.Contains("Both row and column must be specified", exception.Message);
    }

    [Fact]
    public async Task NegativeRow_ShouldThrowException()
    {
        // Arrange
        var arguments = new JsonObject
        {
            ["row"] = -1,
            ["column"] = 0
        };

        // Act & Assert
        var exception = await Assert.ThrowsAsync<ArgumentException>(() => _tool.ExecuteAsync(arguments));
        Assert.Contains("out of range", exception.Message);
    }

    [Fact]
    public async Task NegativeColumn_ShouldThrowException()
    {
        // Arrange
        var arguments = new JsonObject
        {
            ["row"] = 0,
            ["column"] = -1
        };

        // Act & Assert
        var exception = await Assert.ThrowsAsync<ArgumentException>(() => _tool.ExecuteAsync(arguments));
        Assert.Contains("out of range", exception.Message);
    }

    [Fact]
    public async Task RowExceedsMax_ShouldThrowException()
    {
        // Arrange
        var arguments = new JsonObject
        {
            ["row"] = 1048576,
            ["column"] = 0
        };

        // Act & Assert
        var exception = await Assert.ThrowsAsync<ArgumentException>(() => _tool.ExecuteAsync(arguments));
        Assert.Contains("out of range", exception.Message);
    }

    [Fact]
    public async Task ColumnExceedsMax_ShouldThrowException()
    {
        // Arrange
        var arguments = new JsonObject
        {
            ["row"] = 0,
            ["column"] = 16384
        };

        // Act & Assert
        var exception = await Assert.ThrowsAsync<ArgumentException>(() => _tool.ExecuteAsync(arguments));
        Assert.Contains("out of range", exception.Message);
    }

    [Fact]
    public async Task MaxValidRow_ShouldSucceed()
    {
        // Arrange
        var arguments = new JsonObject
        {
            ["row"] = 1048575,
            ["column"] = 0
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.Contains("Row 1048575", result);
    }

    [Fact]
    public async Task MaxValidColumn_ShouldSucceed()
    {
        // Arrange
        var arguments = new JsonObject
        {
            ["row"] = 0,
            ["column"] = 16383
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.Contains("Column 16383", result);
        Assert.Contains("XFD", result);
    }

    [Fact]
    public async Task LowercaseCellAddress_ShouldWork()
    {
        // Arrange
        var arguments = new JsonObject
        {
            ["cellAddress"] = "b2"
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.Contains("Row 1", result);
        Assert.Contains("Column 1", result);
    }
}