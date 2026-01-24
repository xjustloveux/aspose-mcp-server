using AsposeMcpServer.Core.Transport;

namespace AsposeMcpServer.Tests.Integration.Transport;

/// <summary>
///     Integration tests for transport configuration.
/// </summary>
[Trait("Category", "Integration")]
public class TransportConfigTests
{
    /// <summary>
    ///     Verifies that default config values are correct.
    /// </summary>
    [Fact]
    public void Config_Default_HasExpectedValues()
    {
        // Act
        var config = new TransportConfig();

        // Assert
        Assert.Equal(TransportMode.Stdio, config.Mode);
        Assert.Equal(3000, config.Port);
        Assert.Equal("localhost", config.Host);
    }

    /// <summary>
    ///     Verifies that Stdio mode can be set via command line.
    /// </summary>
    [Fact]
    public void Config_StdioArg_SetsStdioMode()
    {
        // Act
        var config = TransportConfig.LoadFromArgs(["--stdio"]);

        // Assert
        Assert.Equal(TransportMode.Stdio, config.Mode);
    }

    /// <summary>
    ///     Verifies that SSE mode can be set via command line.
    /// </summary>
    [Fact]
    public void Config_SseArg_SetsSseMode()
    {
        // Act
        var config = TransportConfig.LoadFromArgs(["--sse"]);

        // Assert
        Assert.Equal(TransportMode.Sse, config.Mode);
    }

    /// <summary>
    ///     Verifies that WebSocket mode can be set via command line.
    /// </summary>
    [Theory]
    [InlineData("--ws")]
    [InlineData("--websocket")]
    public void Config_WebSocketArg_SetsWebSocketMode(string arg)
    {
        // Act
        var config = TransportConfig.LoadFromArgs([arg]);

        // Assert
        Assert.Equal(TransportMode.WebSocket, config.Mode);
    }

    /// <summary>
    ///     Verifies that port can be set via command line in different formats.
    /// </summary>
    [Theory]
    [InlineData(new[] { "--port", "8080" }, 8080)]
    [InlineData(new[] { "--port:8081" }, 8081)]
    [InlineData(new[] { "--port=8082" }, 8082)]
    public void Config_PortArg_SetsPort(string[] args, int expectedPort)
    {
        // Act
        var config = TransportConfig.LoadFromArgs(args);

        // Assert
        Assert.Equal(expectedPort, config.Port);
    }

    /// <summary>
    ///     Verifies that host can be set via command line in different formats.
    /// </summary>
    [Theory]
    [InlineData(new[] { "--host", "0.0.0.0" }, "0.0.0.0")]
    [InlineData(new[] { "--host:127.0.0.1" }, "127.0.0.1")]
    [InlineData(new[] { "--host=192.168.1.1" }, "192.168.1.1")]
    public void Config_HostArg_SetsHost(string[] args, string expectedHost)
    {
        // Act
        var config = TransportConfig.LoadFromArgs(args);

        // Assert
        Assert.Equal(expectedHost, config.Host);
    }

    /// <summary>
    ///     Verifies that validation corrects invalid port values.
    /// </summary>
    [Theory]
    [InlineData(0)]
    [InlineData(-1)]
    [InlineData(65536)]
    [InlineData(100000)]
    public void Config_InvalidPort_ResetsToDefault(int invalidPort)
    {
        // Arrange
        var config = new TransportConfig { Port = invalidPort };

        // Act
        config.Validate();

        // Assert
        Assert.Equal(3000, config.Port);
    }

    /// <summary>
    ///     Verifies that validation corrects invalid host values.
    /// </summary>
    [Theory]
    [InlineData("")]
    [InlineData("   ")]
    [InlineData("invalid-host-name")]
    public void Config_InvalidHost_ResetsToDefault(string invalidHost)
    {
        // Arrange
        var config = new TransportConfig { Host = invalidHost };

        // Act
        config.Validate();

        // Assert
        Assert.Equal("localhost", config.Host);
    }

    /// <summary>
    ///     Verifies that valid hosts pass validation.
    /// </summary>
    [Theory]
    [InlineData("localhost")]
    [InlineData("0.0.0.0")]
    [InlineData("*")]
    [InlineData("127.0.0.1")]
    [InlineData("192.168.1.100")]
    public void Config_ValidHost_RemainsUnchanged(string validHost)
    {
        // Arrange
        var config = new TransportConfig { Host = validHost };

        // Act
        config.Validate();

        // Assert
        Assert.Equal(validHost, config.Host);
    }

    /// <summary>
    ///     Verifies that multiple arguments can be combined.
    /// </summary>
    [Fact]
    public void Config_MultipleArgs_AllApplied()
    {
        // Act
        var config = TransportConfig.LoadFromArgs(["--sse", "--port", "9000", "--host", "0.0.0.0"]);

        // Assert
        Assert.Equal(TransportMode.Sse, config.Mode);
        Assert.Equal(9000, config.Port);
        Assert.Equal("0.0.0.0", config.Host);
    }
}
