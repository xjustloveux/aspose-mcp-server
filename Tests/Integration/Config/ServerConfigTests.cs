using AsposeMcpServer.Core;

namespace AsposeMcpServer.Tests.Integration.Config;

/// <summary>
///     Integration tests for server configuration validation.
/// </summary>
[Trait("Category", "Integration")]
public class ServerConfigTests
{
    #region Invalid Configuration Tests

    /// <summary>
    ///     Verifies that configuration with no tools enabled throws InvalidOperationException.
    /// </summary>
    [Fact]
    public void Config_NoToolsEnabled_ThrowsInvalidOperationException()
    {
        var originalEnv = Environment.GetEnvironmentVariable("ASPOSE_TOOLS");
        try
        {
            Environment.SetEnvironmentVariable("ASPOSE_TOOLS", "none");

            var config = ServerConfig.LoadFromArgs([]);

            Assert.Throws<InvalidOperationException>(() => config.Validate());
        }
        finally
        {
            Environment.SetEnvironmentVariable("ASPOSE_TOOLS", originalEnv);
        }
    }

    #endregion

    #region Valid Configuration Tests

    /// <summary>
    ///     Verifies that valid configuration with all tools enabled passes validation.
    /// </summary>
    [Fact]
    public void Config_AllToolsEnabled_PassesValidation()
    {
        var config = ServerConfig.LoadFromArgs(["--all"]);

        config.Validate();

        Assert.True(config.EnableWord);
        Assert.True(config.EnableExcel);
        Assert.True(config.EnablePowerPoint);
        Assert.True(config.EnablePdf);
    }

    /// <summary>
    ///     Verifies that valid configuration with single tool enabled passes validation.
    /// </summary>
    [Fact]
    public void Config_SingleToolEnabled_PassesValidation()
    {
        var config = ServerConfig.LoadFromArgs(["--word"]);

        config.Validate();

        Assert.True(config.EnableWord);
        Assert.False(config.EnableExcel);
    }

    /// <summary>
    ///     Verifies that valid configuration with multiple tools enabled passes validation.
    /// </summary>
    [Fact]
    public void Config_MultipleToolsEnabled_PassesValidation()
    {
        var config = ServerConfig.LoadFromArgs(["--word", "--excel"]);

        config.Validate();

        Assert.True(config.EnableWord);
        Assert.True(config.EnableExcel);
        Assert.False(config.EnablePowerPoint);
        Assert.False(config.EnablePdf);
    }

    #endregion

    #region Environment Variable Override Tests

    /// <summary>
    ///     Verifies that command line arguments override environment variables.
    /// </summary>
    [Fact]
    public void Config_CommandLineOverridesEnvironment()
    {
        var originalEnv = Environment.GetEnvironmentVariable("ASPOSE_TOOLS");
        try
        {
            Environment.SetEnvironmentVariable("ASPOSE_TOOLS", "word");

            var config = ServerConfig.LoadFromArgs(["--excel"]);

            Assert.False(config.EnableWord);
            Assert.True(config.EnableExcel);
        }
        finally
        {
            Environment.SetEnvironmentVariable("ASPOSE_TOOLS", originalEnv);
        }
    }

    /// <summary>
    ///     Verifies that environment variables are loaded when no command line args.
    /// </summary>
    [Fact]
    public void Config_EnvironmentVariableLoaded()
    {
        var originalEnv = Environment.GetEnvironmentVariable("ASPOSE_TOOLS");
        try
        {
            Environment.SetEnvironmentVariable("ASPOSE_TOOLS", "pdf");

            var config = ServerConfig.LoadFromArgs([]);

            Assert.True(config.EnablePdf);
            Assert.False(config.EnableWord);
        }
        finally
        {
            Environment.SetEnvironmentVariable("ASPOSE_TOOLS", originalEnv);
        }
    }

    #endregion

    #region GetEnabledToolsInfo Tests

    /// <summary>
    ///     Verifies that GetEnabledToolsInfo returns correct information.
    /// </summary>
    [Fact]
    public void Config_GetEnabledToolsInfo_ReturnsCorrectString()
    {
        var config = ServerConfig.LoadFromArgs(["--word", "--excel"]);

        var info = config.GetEnabledToolsInfo();

        Assert.Contains("Word", info);
        Assert.Contains("Excel", info);
        Assert.DoesNotContain("PowerPoint", info);
        Assert.DoesNotContain("PDF", info);
    }

    /// <summary>
    ///     Verifies that GetEnabledToolsInfo returns "None" when no tools enabled.
    /// </summary>
    [Fact]
    public void Config_GetEnabledToolsInfo_NoTools_ReturnsNone()
    {
        var originalEnv = Environment.GetEnvironmentVariable("ASPOSE_TOOLS");
        try
        {
            Environment.SetEnvironmentVariable("ASPOSE_TOOLS", "none");

            var config = ServerConfig.LoadFromArgs([]);

            var info = config.GetEnabledToolsInfo();

            Assert.Equal("None", info);
        }
        finally
        {
            Environment.SetEnvironmentVariable("ASPOSE_TOOLS", originalEnv);
        }
    }

    #endregion
}
