using AsposeMcpServer.Core;

namespace AsposeMcpServer.Tests.Core;

/// <summary>
///     Unit tests for LicenseManager class
/// </summary>
public class LicenseManagerTests : IDisposable
{
    private readonly StringWriter _consoleError;
    private readonly StringWriter _consoleOut;
    private readonly TextWriter _originalConsoleError;
    private readonly TextWriter _originalConsoleOut;

    public LicenseManagerTests()
    {
        _originalConsoleOut = Console.Out;
        _originalConsoleError = Console.Error;
        _consoleOut = new StringWriter();
        _consoleError = new StringWriter();
        Console.SetOut(_consoleOut);
        Console.SetError(_consoleError);
    }

    public void Dispose()
    {
        Console.SetOut(_originalConsoleOut);
        Console.SetError(_originalConsoleError);
        _consoleOut.Dispose();
        _consoleError.Dispose();
    }

    #region SetLicense Tests

    [Fact]
    public void SetLicense_WithNoLicenseFile_ShouldOutputMessage()
    {
        // Use LoadFromArgs with a nonexistent license path
        var config = ServerConfig.LoadFromArgs(["--all", "--license:nonexistent_license.lic"]);

        LicenseManager.SetLicense(config);

        var errorOutput = _consoleError.ToString();
        // Either license loaded successfully or shows warning about no license file
        Assert.True(
            errorOutput.Contains("license", StringComparison.OrdinalIgnoreCase),
            "Expected output to mention 'license'");
    }

    [Fact]
    public void SetLicense_WithDefaultConfig_ShouldNotThrow()
    {
        var config = ServerConfig.LoadFromArgs([]);

        var exception = Record.Exception(() => LicenseManager.SetLicense(config));

        Assert.Null(exception);
    }

    [Fact]
    public void SetLicense_WithAllComponentsEnabled_ShouldSearchForLicenses()
    {
        var config = ServerConfig.LoadFromArgs(["--all"]);

        var exception = Record.Exception(() => LicenseManager.SetLicense(config));

        // Should not throw even without a license file
        Assert.Null(exception);
    }

    [Fact]
    public void SetLicense_WithSpecificLicensePath_ShouldSearchAndOutputResult()
    {
        var config = ServerConfig.LoadFromArgs(["--all", "--license:custom_path/license.lic"]);

        LicenseManager.SetLicense(config);

        var errorOutput = _consoleError.ToString();
        // Should output something about license - either found or not found
        Assert.True(
            errorOutput.Contains("license", StringComparison.OrdinalIgnoreCase),
            "Expected output to mention 'license'");
    }

    [Fact]
    public void SetLicense_WithWordOnly_ShouldNotThrow()
    {
        var config = ServerConfig.LoadFromArgs(["--word"]);

        var exception = Record.Exception(() => LicenseManager.SetLicense(config));

        Assert.Null(exception);
    }

    [Fact]
    public void SetLicense_WithExcelOnly_ShouldNotThrow()
    {
        var config = ServerConfig.LoadFromArgs(["--excel"]);

        var exception = Record.Exception(() => LicenseManager.SetLicense(config));

        Assert.Null(exception);
    }

    [Fact]
    public void SetLicense_WithPowerPointOnly_ShouldNotThrow()
    {
        var config = ServerConfig.LoadFromArgs(["--powerpoint"]);

        var exception = Record.Exception(() => LicenseManager.SetLicense(config));

        Assert.Null(exception);
    }

    [Fact]
    public void SetLicense_WithPdfOnly_ShouldNotThrow()
    {
        var config = ServerConfig.LoadFromArgs(["--pdf"]);

        var exception = Record.Exception(() => LicenseManager.SetLicense(config));

        Assert.Null(exception);
    }

    #endregion

    #region Console Output Tests

    [Fact]
    public void SetLicense_ShouldRestoreConsoleOut()
    {
        var config = ServerConfig.LoadFromArgs(["--word"]);

        // Reset console to capture output properly
        Console.SetOut(_originalConsoleOut);
        Console.SetError(_consoleError);

        LicenseManager.SetLicense(config);

        // Console.Out should be restored after SetLicense
        Assert.Same(_originalConsoleOut, Console.Out);
    }

    [Fact]
    public void SetLicense_ShouldOutputEvaluationModeMessage()
    {
        var config = ServerConfig.LoadFromArgs(["--all"]);

        LicenseManager.SetLicense(config);

        var errorOutput = _consoleError.ToString();
        // Should mention either license loaded or evaluation mode
        Assert.True(
            errorOutput.Contains("license", StringComparison.OrdinalIgnoreCase) ||
            errorOutput.Contains("evaluation", StringComparison.OrdinalIgnoreCase));
    }

    #endregion
}