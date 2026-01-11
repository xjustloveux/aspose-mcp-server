using AsposeMcpServer.Core;

namespace AsposeMcpServer.Tests.Core;

/// <summary>
///     Unit tests for ServerConfig class
/// </summary>
public class ServerConfigTests
{
    #region Default Values Tests

    [Fact]
    public void LoadFromArgs_NoArgs_ShouldHaveDefaultValues()
    {
        var config = ServerConfig.LoadFromArgs([]);

        Assert.True(config.EnableWord);
        Assert.True(config.EnableExcel);
        Assert.True(config.EnablePowerPoint);
        Assert.True(config.EnablePdf);
    }

    #endregion

    #region Tool Enable Tests

    [Fact]
    public void LoadFromArgs_WithWord_ShouldEnableOnlyWord()
    {
        var config = ServerConfig.LoadFromArgs(["--word"]);

        Assert.True(config.EnableWord);
        Assert.False(config.EnableExcel);
        Assert.False(config.EnablePowerPoint);
        Assert.False(config.EnablePdf);
    }

    [Fact]
    public void LoadFromArgs_WithExcel_ShouldEnableOnlyExcel()
    {
        var config = ServerConfig.LoadFromArgs(["--excel"]);

        Assert.False(config.EnableWord);
        Assert.True(config.EnableExcel);
        Assert.False(config.EnablePowerPoint);
        Assert.False(config.EnablePdf);
    }

    [Fact]
    public void LoadFromArgs_WithPowerPoint_ShouldEnableOnlyPowerPoint()
    {
        var config = ServerConfig.LoadFromArgs(["--powerpoint"]);

        Assert.False(config.EnableWord);
        Assert.False(config.EnableExcel);
        Assert.True(config.EnablePowerPoint);
        Assert.False(config.EnablePdf);
    }

    [Fact]
    public void LoadFromArgs_WithPpt_ShouldEnableOnlyPowerPoint()
    {
        var config = ServerConfig.LoadFromArgs(["--ppt"]);

        Assert.False(config.EnableWord);
        Assert.False(config.EnableExcel);
        Assert.True(config.EnablePowerPoint);
        Assert.False(config.EnablePdf);
    }

    [Fact]
    public void LoadFromArgs_WithPdf_ShouldEnableOnlyPdf()
    {
        var config = ServerConfig.LoadFromArgs(["--pdf"]);

        Assert.False(config.EnableWord);
        Assert.False(config.EnableExcel);
        Assert.False(config.EnablePowerPoint);
        Assert.True(config.EnablePdf);
    }

    [Fact]
    public void LoadFromArgs_WithAll_ShouldEnableAll()
    {
        var config = ServerConfig.LoadFromArgs(["--all"]);

        Assert.True(config.EnableWord);
        Assert.True(config.EnableExcel);
        Assert.True(config.EnablePowerPoint);
        Assert.True(config.EnablePdf);
    }

    [Fact]
    public void LoadFromArgs_WithMultipleTools_ShouldEnableSelected()
    {
        var config = ServerConfig.LoadFromArgs(["--word", "--pdf"]);

        Assert.True(config.EnableWord);
        Assert.False(config.EnableExcel);
        Assert.False(config.EnablePowerPoint);
        Assert.True(config.EnablePdf);
    }

    #endregion

    #region License Tests

    [Fact]
    public void LoadFromArgs_WithLicenseColon_ShouldSetLicensePath()
    {
        var config = ServerConfig.LoadFromArgs(["--license:Aspose.Total.lic"]);

        Assert.Equal("Aspose.Total.lic", config.LicensePath);
    }

    [Fact]
    public void LoadFromArgs_WithLicenseEquals_ShouldSetLicensePath()
    {
        var config = ServerConfig.LoadFromArgs(["--license=/path/to/license.lic"]);

        Assert.Equal("/path/to/license.lic", config.LicensePath);
    }

    [Fact]
    public void LoadFromArgs_WithEnvironmentVariable_ShouldSetLicensePath()
    {
        Environment.SetEnvironmentVariable("ASPOSE_LICENSE_PATH", "/env/license.lic");
        try
        {
            var config = ServerConfig.LoadFromArgs([]);

            Assert.Equal("/env/license.lic", config.LicensePath);
        }
        finally
        {
            Environment.SetEnvironmentVariable("ASPOSE_LICENSE_PATH", null);
        }
    }

    #endregion

    #region GetEnabledToolsInfo Tests

    [Fact]
    public void GetEnabledToolsInfo_AllEnabled_ShouldReturnAll()
    {
        var config = ServerConfig.LoadFromArgs(["--all"]);

        var info = config.GetEnabledToolsInfo();

        Assert.Contains("Word", info);
        Assert.Contains("Excel", info);
        Assert.Contains("PowerPoint", info);
        Assert.Contains("PDF", info);
    }

    [Fact]
    public void GetEnabledToolsInfo_SomeEnabled_ShouldReturnEnabled()
    {
        var config = ServerConfig.LoadFromArgs(["--word", "--pdf"]);

        var info = config.GetEnabledToolsInfo();

        Assert.Contains("Word", info);
        Assert.Contains("PDF", info);
        Assert.DoesNotContain("Excel", info);
        Assert.DoesNotContain("PowerPoint", info);
    }

    [Fact]
    public void GetEnabledToolsInfo_NoneEnabled_ShouldReturnNone()
    {
        var originalValue = Environment.GetEnvironmentVariable("ASPOSE_TOOLS");
        try
        {
            Environment.SetEnvironmentVariable("ASPOSE_TOOLS", "invalid");
            var config = ServerConfig.LoadFromArgs([]);

            var info = config.GetEnabledToolsInfo();

            Assert.Equal("None", info);
        }
        finally
        {
            Environment.SetEnvironmentVariable("ASPOSE_TOOLS", originalValue);
        }
    }

    #endregion

    #region Validate Tests

    [Fact]
    public void Validate_AllEnabled_ShouldNotThrow()
    {
        var config = ServerConfig.LoadFromArgs([]);

        var ex = Record.Exception(() => config.Validate());

        Assert.Null(ex);
    }

    [Fact]
    public void Validate_SomeEnabled_ShouldNotThrow()
    {
        var config = ServerConfig.LoadFromArgs(["--word"]);

        var ex = Record.Exception(() => config.Validate());

        Assert.Null(ex);
    }

    [Fact]
    public void Validate_NoneEnabled_ShouldThrow()
    {
        var originalValue = Environment.GetEnvironmentVariable("ASPOSE_TOOLS");
        try
        {
            Environment.SetEnvironmentVariable("ASPOSE_TOOLS", "invalid");
            var config = ServerConfig.LoadFromArgs([]);

            var ex = Assert.Throws<InvalidOperationException>(() => config.Validate());

            Assert.Contains("At least one tool category must be enabled", ex.Message);
        }
        finally
        {
            Environment.SetEnvironmentVariable("ASPOSE_TOOLS", originalValue);
        }
    }

    #endregion

    #region Case Insensitivity Tests

    [Fact]
    public void LoadFromArgs_UpperCaseArgs_ShouldWork()
    {
        var config = ServerConfig.LoadFromArgs(["--WORD", "--PDF"]);

        Assert.True(config.EnableWord);
        Assert.True(config.EnablePdf);
    }

    [Fact]
    public void LoadFromArgs_MixedCaseArgs_ShouldWork()
    {
        var config = ServerConfig.LoadFromArgs(["--Word", "--Excel"]);

        Assert.True(config.EnableWord);
        Assert.True(config.EnableExcel);
    }

    #endregion
}
