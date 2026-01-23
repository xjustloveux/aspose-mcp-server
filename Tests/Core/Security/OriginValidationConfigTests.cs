using AsposeMcpServer.Core.Security;

namespace AsposeMcpServer.Tests.Core.Security;

public class OriginValidationConfigTests
{
    [Fact]
    public void DefaultConfig_ShouldHaveExpectedDefaults()
    {
        var config = new OriginValidationConfig();

        Assert.True(config.Enabled);
        Assert.True(config.AllowLocalhost);
        Assert.True(config.AllowMissingOrigin);
        Assert.Null(config.AllowedOrigins);
        Assert.Equal(["/health", "/ready"], config.ExcludedPaths);
    }

    [Fact]
    public void LoadFromArgs_WithNoOriginValidation_ShouldDisable()
    {
        var args = new[] { "--no-origin-validation" };

        var config = OriginValidationConfig.LoadFromArgs(args);

        Assert.False(config.Enabled);
    }

    [Fact]
    public void LoadFromArgs_WithNoLocalhost_ShouldDisableLocalhost()
    {
        var args = new[] { "--no-localhost" };

        var config = OriginValidationConfig.LoadFromArgs(args);

        Assert.False(config.AllowLocalhost);
    }

    [Fact]
    public void LoadFromArgs_WithRequireOrigin_ShouldDisallowMissingOrigin()
    {
        var args = new[] { "--require-origin" };

        var config = OriginValidationConfig.LoadFromArgs(args);

        Assert.False(config.AllowMissingOrigin);
    }

    [Fact]
    public void LoadFromArgs_WithAllowedOrigins_ShouldParseOrigins()
    {
        var args = new[] { "--allowed-origins:https://example.com,https://app.example.com" };

        var config = OriginValidationConfig.LoadFromArgs(args);

        Assert.NotNull(config.AllowedOrigins);
        Assert.Equal(2, config.AllowedOrigins.Length);
        Assert.Contains("https://example.com", config.AllowedOrigins);
        Assert.Contains("https://app.example.com", config.AllowedOrigins);
    }

    [Fact]
    public void LoadFromArgs_WithEmptyArgs_ShouldReturnDefaults()
    {
        var args = Array.Empty<string>();

        var config = OriginValidationConfig.LoadFromArgs(args);

        Assert.True(config.Enabled);
        Assert.True(config.AllowLocalhost);
        Assert.True(config.AllowMissingOrigin);
    }

    [Fact]
    public void LoadFromArgs_WithMultipleFlags_ShouldApplyAll()
    {
        var args = new[]
        {
            "--no-localhost",
            "--require-origin",
            "--allowed-origins:https://trusted.com"
        };

        var config = OriginValidationConfig.LoadFromArgs(args);

        Assert.True(config.Enabled);
        Assert.False(config.AllowLocalhost);
        Assert.False(config.AllowMissingOrigin);
        Assert.NotNull(config.AllowedOrigins);
        Assert.Single(config.AllowedOrigins);
    }
}
