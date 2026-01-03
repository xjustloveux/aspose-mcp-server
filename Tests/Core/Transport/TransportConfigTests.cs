using AsposeMcpServer.Core.Transport;

namespace AsposeMcpServer.Tests.Core.Transport;

/// <summary>
///     Unit tests for TransportConfig class
/// </summary>
public class TransportConfigTests
{
    #region Default Values Tests

    [Fact]
    public void LoadFromArgs_NoArgs_ShouldHaveDefaultValues()
    {
        var config = TransportConfig.LoadFromArgs([]);

        Assert.Equal(TransportMode.Stdio, config.Mode);
        Assert.Equal(3000, config.Port);
        Assert.Equal("localhost", config.Host);
    }

    #endregion

    #region Transport Mode Tests

    [Fact]
    public void LoadFromArgs_WithStdio_ShouldSetStdioMode()
    {
        var config = TransportConfig.LoadFromArgs(["--stdio"]);

        Assert.Equal(TransportMode.Stdio, config.Mode);
    }

    [Fact]
    public void LoadFromArgs_WithSse_ShouldSetSseMode()
    {
        var config = TransportConfig.LoadFromArgs(["--sse"]);

        Assert.Equal(TransportMode.Sse, config.Mode);
    }

    [Fact]
    public void LoadFromArgs_WithWebSocket_ShouldSetWebSocketMode()
    {
        var config = TransportConfig.LoadFromArgs(["--websocket"]);

        Assert.Equal(TransportMode.WebSocket, config.Mode);
    }

    [Fact]
    public void LoadFromArgs_WithWs_ShouldSetWebSocketMode()
    {
        var config = TransportConfig.LoadFromArgs(["--ws"]);

        Assert.Equal(TransportMode.WebSocket, config.Mode);
    }

    #endregion

    #region Port Tests

    [Fact]
    public void LoadFromArgs_WithPortArg_ShouldSetPort()
    {
        var config = TransportConfig.LoadFromArgs(["--port", "8080"]);

        Assert.Equal(8080, config.Port);
    }

    [Fact]
    public void LoadFromArgs_WithPortColon_ShouldSetPort()
    {
        var config = TransportConfig.LoadFromArgs(["--port:8080"]);

        Assert.Equal(8080, config.Port);
    }

    [Fact]
    public void LoadFromArgs_WithPortEquals_ShouldSetPort()
    {
        var config = TransportConfig.LoadFromArgs(["--port=8080"]);

        Assert.Equal(8080, config.Port);
    }

    [Fact]
    public void LoadFromArgs_WithInvalidPort_ShouldKeepDefault()
    {
        var config = TransportConfig.LoadFromArgs(["--port", "invalid"]);

        Assert.Equal(3000, config.Port);
    }

    #endregion

    #region Host Tests

    [Fact]
    public void LoadFromArgs_WithHostArg_ShouldSetHost()
    {
        var config = TransportConfig.LoadFromArgs(["--host", "0.0.0.0"]);

        Assert.Equal("0.0.0.0", config.Host);
    }

    [Fact]
    public void LoadFromArgs_WithHostColon_ShouldSetHost()
    {
        var config = TransportConfig.LoadFromArgs(["--host:0.0.0.0"]);

        Assert.Equal("0.0.0.0", config.Host);
    }

    [Fact]
    public void LoadFromArgs_WithHostEquals_ShouldSetHost()
    {
        var config = TransportConfig.LoadFromArgs(["--host=192.168.1.1"]);

        Assert.Equal("192.168.1.1", config.Host);
    }

    #endregion

    #region Environment Variable Tests

    [Fact]
    public void LoadFromArgs_WithTransportEnvVar_ShouldSetMode()
    {
        Environment.SetEnvironmentVariable("ASPOSE_TRANSPORT", "sse");
        try
        {
            var config = TransportConfig.LoadFromArgs([]);

            Assert.Equal(TransportMode.Sse, config.Mode);
        }
        finally
        {
            Environment.SetEnvironmentVariable("ASPOSE_TRANSPORT", null);
        }
    }

    [Fact]
    public void LoadFromArgs_WithPortEnvVar_ShouldSetPort()
    {
        Environment.SetEnvironmentVariable("ASPOSE_PORT", "9000");
        try
        {
            var config = TransportConfig.LoadFromArgs([]);

            Assert.Equal(9000, config.Port);
        }
        finally
        {
            Environment.SetEnvironmentVariable("ASPOSE_PORT", null);
        }
    }

    [Fact]
    public void LoadFromArgs_WithHostEnvVar_ShouldSetHost()
    {
        Environment.SetEnvironmentVariable("ASPOSE_HOST", "0.0.0.0");
        try
        {
            var config = TransportConfig.LoadFromArgs([]);

            Assert.Equal("0.0.0.0", config.Host);
        }
        finally
        {
            Environment.SetEnvironmentVariable("ASPOSE_HOST", null);
        }
    }

    [Fact]
    public void LoadFromArgs_CommandLineOverridesEnvVar()
    {
        Environment.SetEnvironmentVariable("ASPOSE_PORT", "9000");
        try
        {
            var config = TransportConfig.LoadFromArgs(["--port:8080"]);

            Assert.Equal(8080, config.Port);
        }
        finally
        {
            Environment.SetEnvironmentVariable("ASPOSE_PORT", null);
        }
    }

    #endregion

    #region Case Insensitivity Tests

    [Fact]
    public void LoadFromArgs_UpperCaseArgs_ShouldWork()
    {
        var config = TransportConfig.LoadFromArgs(["--SSE"]);

        Assert.Equal(TransportMode.Sse, config.Mode);
    }

    [Fact]
    public void LoadFromArgs_MixedCaseArgs_ShouldWork()
    {
        var config = TransportConfig.LoadFromArgs(["--WebSocket"]);

        Assert.Equal(TransportMode.WebSocket, config.Mode);
    }

    #endregion
}