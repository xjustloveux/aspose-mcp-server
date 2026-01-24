using AsposeMcpServer.Core.Transport;

namespace AsposeMcpServer.Tests.Integration.Transport;

/// <summary>
///     Integration tests for SSE transport configuration.
/// </summary>
[Trait("Category", "Integration")]
public class SseTransportTests
{
    /// <summary>
    ///     Verifies that SSE mode is correctly parsed from command line.
    /// </summary>
    [Fact]
    public void Sse_ModeFromArgs_ParsedCorrectly()
    {
        var config = TransportConfig.LoadFromArgs(["--sse"]);

        Assert.Equal(TransportMode.Sse, config.Mode);
    }

    /// <summary>
    ///     Verifies that SSE mode with port is configured correctly.
    /// </summary>
    [Fact]
    public void Sse_WithPort_ConfiguredCorrectly()
    {
        var config = TransportConfig.LoadFromArgs(["--sse", "--port", "8080"]);

        Assert.Equal(TransportMode.Sse, config.Mode);
        Assert.Equal(8080, config.Port);
    }

    /// <summary>
    ///     Verifies that SSE mode with host is configured correctly.
    /// </summary>
    [Fact]
    public void Sse_WithHost_ConfiguredCorrectly()
    {
        var config = TransportConfig.LoadFromArgs(["--sse", "--host", "0.0.0.0"]);

        Assert.Equal(TransportMode.Sse, config.Mode);
        Assert.Equal("0.0.0.0", config.Host);
    }

    /// <summary>
    ///     Verifies that SSE mode with all options is configured correctly.
    /// </summary>
    [Fact]
    public void Sse_WithAllOptions_ConfiguredCorrectly()
    {
        var config = TransportConfig.LoadFromArgs(["--sse", "--port", "9090", "--host", "127.0.0.1"]);

        Assert.Equal(TransportMode.Sse, config.Mode);
        Assert.Equal(9090, config.Port);
        Assert.Equal("127.0.0.1", config.Host);
    }

    /// <summary>
    ///     Verifies that SSE mode validates configuration.
    /// </summary>
    [Fact]
    public void Sse_InvalidConfig_ValidatesCorrectly()
    {
        var config = TransportConfig.LoadFromArgs(["--sse", "--port", "-1"]);

        config.Validate();

        Assert.Equal(TransportMode.Sse, config.Mode);
        Assert.Equal(3000, config.Port);
    }
}
