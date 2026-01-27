using AsposeMcpServer.Core.Transport;

namespace AsposeMcpServer.Tests.Integration.Transport;

/// <summary>
///     Integration tests for HTTP (Streamable HTTP) transport configuration.
/// </summary>
[Trait("Category", "Integration")]
public class HttpTransportTests
{
    /// <summary>
    ///     Verifies that HTTP mode is correctly parsed from command line.
    /// </summary>
    [Fact]
    public void Http_ModeFromArgs_ParsedCorrectly()
    {
        var config = TransportConfig.LoadFromArgs(["--http"]);

        Assert.Equal(TransportMode.Http, config.Mode);
    }

    /// <summary>
    ///     Verifies that HTTP mode with port is configured correctly.
    /// </summary>
    [Fact]
    public void Http_WithPort_ConfiguredCorrectly()
    {
        var config = TransportConfig.LoadFromArgs(["--http", "--port", "8080"]);

        Assert.Equal(TransportMode.Http, config.Mode);
        Assert.Equal(8080, config.Port);
    }

    /// <summary>
    ///     Verifies that HTTP mode with host is configured correctly.
    /// </summary>
    [Fact]
    public void Http_WithHost_ConfiguredCorrectly()
    {
        var config = TransportConfig.LoadFromArgs(["--http", "--host", "0.0.0.0"]);

        Assert.Equal(TransportMode.Http, config.Mode);
        Assert.Equal("0.0.0.0", config.Host);
    }

    /// <summary>
    ///     Verifies that HTTP mode with all options is configured correctly.
    /// </summary>
    [Fact]
    public void Http_WithAllOptions_ConfiguredCorrectly()
    {
        var config = TransportConfig.LoadFromArgs(["--http", "--port", "9090", "--host", "127.0.0.1"]);

        Assert.Equal(TransportMode.Http, config.Mode);
        Assert.Equal(9090, config.Port);
        Assert.Equal("127.0.0.1", config.Host);
    }

    /// <summary>
    ///     Verifies that HTTP mode validates configuration.
    /// </summary>
    [Fact]
    public void Http_InvalidConfig_ValidatesCorrectly()
    {
        var config = TransportConfig.LoadFromArgs(["--http", "--port", "-1"]);

        config.Validate();

        Assert.Equal(TransportMode.Http, config.Mode);
        Assert.Equal(3000, config.Port);
    }
}
