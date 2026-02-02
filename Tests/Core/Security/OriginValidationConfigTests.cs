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

    [Fact]
    public void LoadFromArgs_WithEqualsFormat_ParsesCorrectly()
    {
        var args = new[] { "--allowed-origins=https://a.com,https://b.com" };

        var config = OriginValidationConfig.LoadFromArgs(args);

        Assert.NotNull(config.AllowedOrigins);
        Assert.Equal(2, config.AllowedOrigins.Length);
        Assert.Contains("https://a.com", config.AllowedOrigins);
        Assert.Contains("https://b.com", config.AllowedOrigins);
    }

    [Fact]
    public void LoadFromArgs_WithSeparateArgFormat_ParsesCorrectly()
    {
        var args = new[] { "--allowed-origins", "https://x.com,https://y.com" };

        var config = OriginValidationConfig.LoadFromArgs(args);

        Assert.NotNull(config.AllowedOrigins);
        Assert.Equal(2, config.AllowedOrigins.Length);
        Assert.Contains("https://x.com", config.AllowedOrigins);
        Assert.Contains("https://y.com", config.AllowedOrigins);
    }

    [Fact]
    public void LoadFromArgs_WithEnvironmentVariables_ParsesAllOptions()
    {
        var envVars = new Dictionary<string, string?>
        {
            { "ASPOSE_ORIGIN_VALIDATION", "false" },
            { "ASPOSE_ALLOW_LOCALHOST", "false" },
            { "ASPOSE_ALLOW_MISSING_ORIGIN", "false" },
            { "ASPOSE_ALLOWED_ORIGINS", "https://env1.com,https://env2.com" }
        };

        try
        {
            foreach (var kv in envVars)
                Environment.SetEnvironmentVariable(kv.Key, kv.Value);

            var config = OriginValidationConfig.LoadFromArgs([]);

            Assert.False(config.Enabled);
            Assert.False(config.AllowLocalhost);
            Assert.False(config.AllowMissingOrigin);
            Assert.NotNull(config.AllowedOrigins);
            Assert.Equal(2, config.AllowedOrigins.Length);
        }
        finally
        {
            foreach (var kv in envVars)
                Environment.SetEnvironmentVariable(kv.Key, null);
        }
    }

    [Fact]
    public void LoadFromArgs_ArgsOverrideEnvironmentVariables()
    {
        try
        {
            Environment.SetEnvironmentVariable("ASPOSE_ALLOWED_ORIGINS", "https://env.com");

            var args = new[] { "--allowed-origins:https://arg.com" };
            var config = OriginValidationConfig.LoadFromArgs(args);

            Assert.NotNull(config.AllowedOrigins);
            Assert.Single(config.AllowedOrigins);
            Assert.Contains("https://arg.com", config.AllowedOrigins);
        }
        finally
        {
            Environment.SetEnvironmentVariable("ASPOSE_ALLOWED_ORIGINS", null);
        }
    }
}
