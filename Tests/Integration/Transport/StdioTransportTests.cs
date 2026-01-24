using AsposeMcpServer.Core.Transport;

namespace AsposeMcpServer.Tests.Integration.Transport;

/// <summary>
///     Integration tests for Stdio transport configuration.
/// </summary>
[Trait("Category", "Integration")]
public class StdioTransportTests
{
    /// <summary>
    ///     Verifies that Stdio is the default transport mode.
    /// </summary>
    [Fact]
    public void Stdio_IsDefaultMode()
    {
        var config = new TransportConfig();

        Assert.Equal(TransportMode.Stdio, config.Mode);
    }

    /// <summary>
    ///     Verifies that Stdio mode is correctly parsed from command line.
    /// </summary>
    [Fact]
    public void Stdio_ModeFromArgs_ParsedCorrectly()
    {
        var config = TransportConfig.LoadFromArgs(["--stdio"]);

        Assert.Equal(TransportMode.Stdio, config.Mode);
    }

    /// <summary>
    ///     Verifies that Stdio mode ignores port setting.
    /// </summary>
    [Fact]
    public void Stdio_WithPort_PortStillSet()
    {
        var config = TransportConfig.LoadFromArgs(["--stdio", "--port", "8080"]);

        Assert.Equal(TransportMode.Stdio, config.Mode);
        Assert.Equal(8080, config.Port);
    }

    /// <summary>
    ///     Verifies that Stdio mode ignores host setting.
    /// </summary>
    [Fact]
    public void Stdio_WithHost_HostStillSet()
    {
        var config = TransportConfig.LoadFromArgs(["--stdio", "--host", "0.0.0.0"]);

        Assert.Equal(TransportMode.Stdio, config.Mode);
        Assert.Equal("0.0.0.0", config.Host);
    }

    /// <summary>
    ///     Verifies that Stdio mode can be combined with other tool arguments.
    /// </summary>
    [Fact]
    public void Stdio_WithOtherArgs_ParsedCorrectly()
    {
        var config = TransportConfig.LoadFromArgs(["--stdio", "--word", "--excel"]);

        Assert.Equal(TransportMode.Stdio, config.Mode);
    }

    /// <summary>
    ///     Verifies that mode precedence works correctly (last wins).
    /// </summary>
    [Fact]
    public void Stdio_AfterSse_StdioWins()
    {
        var config = TransportConfig.LoadFromArgs(["--sse", "--stdio"]);

        Assert.Equal(TransportMode.Stdio, config.Mode);
    }
}
