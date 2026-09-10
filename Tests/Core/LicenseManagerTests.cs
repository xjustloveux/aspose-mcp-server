using AsposeMcpServer.Core;

namespace AsposeMcpServer.Tests.Core;

/// <summary>
///     Unit tests for LicenseManager class
/// </summary>
/// <remarks>
///     Licensed-only, by necessity rather than preference. Every test here calls
///     <see cref="LicenseManager.SetLicense" />, and an Aspose licence applies to the whole
///     process — so running one of them in a host asked to stay unlicensed licenses that host, and
///     every test scheduled afterwards silently stops testing evaluation behaviour (R13-T01).
///     There is no way to exercise licence loading without loading a licence, so in that host they
///     skip. <see cref="Infrastructure.EvaluationHostTests" /> is what notices if they ever stop.
/// </remarks>
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

    /// <summary>Refuses to run when this host was asked not to load a licence.</summary>
    private static void SkipWhenTheHostMustStayUnlicensed()
    {
        var skip = Environment.GetEnvironmentVariable("SKIP_ASPOSE_LICENSE");
        Skip.If(
            string.Equals(skip, "true", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(skip, "1", StringComparison.OrdinalIgnoreCase),
            "Loading a licence here would license the whole process and make every later "
            + "evaluation-mode test meaningless (R13-T01).");
    }

    #region SetLicense Tests

    [SkippableFact]
    public void SetLicense_WithNoLicenseFile_ShouldOutputMessage()
    {
        SkipWhenTheHostMustStayUnlicensed();

        var config = ServerConfig.LoadFromArgs(["--all", "--license:nonexistent_license.lic"]);

        LicenseManager.SetLicense(config);

        var errorOutput = _consoleError.ToString();
        // Either license loaded successfully or shows warning about no license file
        Assert.True(
            errorOutput.Contains("license", StringComparison.OrdinalIgnoreCase),
            "Expected output to mention 'license'");
    }

    [SkippableFact]
    public void SetLicense_WithDefaultConfig_ShouldNotThrowAndRestoreConsole()
    {
        SkipWhenTheHostMustStayUnlicensed();

        var config = ServerConfig.LoadFromArgs([]);

        var exception = Record.Exception(() => LicenseManager.SetLicense(config));

        Assert.Null(exception);
        Assert.NotEqual(TextWriter.Null, Console.Out);
    }

    [SkippableFact]
    public void SetLicense_WithAllComponentsEnabled_ShouldSearchForLicensesAndOutputResult()
    {
        SkipWhenTheHostMustStayUnlicensed();

        var config = ServerConfig.LoadFromArgs(["--all"]);

        var exception = Record.Exception(() => LicenseManager.SetLicense(config));

        Assert.Null(exception);
        var errorOutput = _consoleError.ToString();
        Assert.False(string.IsNullOrEmpty(errorOutput), "Expected license status output on stderr");
    }

    [SkippableFact]
    public void SetLicense_WithSpecificLicensePath_ShouldSearchAndOutputResult()
    {
        SkipWhenTheHostMustStayUnlicensed();

        var config = ServerConfig.LoadFromArgs(["--all", "--license:custom_path/license.lic"]);

        LicenseManager.SetLicense(config);

        var errorOutput = _consoleError.ToString();
        Assert.True(
            errorOutput.Contains("license", StringComparison.OrdinalIgnoreCase),
            "Expected output to mention 'license'");
    }

    [SkippableFact]
    public void SetLicense_WithWordOnly_ShouldOutputLicenseStatus()
    {
        SkipWhenTheHostMustStayUnlicensed();

        var config = ServerConfig.LoadFromArgs(["--word"]);

        var exception = Record.Exception(() => LicenseManager.SetLicense(config));

        Assert.Null(exception);
        var errorOutput = _consoleError.ToString();
        Assert.False(string.IsNullOrEmpty(errorOutput), "Expected license status output on stderr");
    }

    [SkippableFact]
    public void SetLicense_WithExcelOnly_ShouldOutputLicenseStatus()
    {
        SkipWhenTheHostMustStayUnlicensed();

        var config = ServerConfig.LoadFromArgs(["--excel"]);

        var exception = Record.Exception(() => LicenseManager.SetLicense(config));

        Assert.Null(exception);
        var errorOutput = _consoleError.ToString();
        Assert.False(string.IsNullOrEmpty(errorOutput), "Expected license status output on stderr");
    }

    [SkippableFact]
    public void SetLicense_WithPowerPointOnly_ShouldOutputLicenseStatus()
    {
        SkipWhenTheHostMustStayUnlicensed();

        var config = ServerConfig.LoadFromArgs(["--powerpoint"]);

        var exception = Record.Exception(() => LicenseManager.SetLicense(config));

        Assert.Null(exception);
        var errorOutput = _consoleError.ToString();
        Assert.False(string.IsNullOrEmpty(errorOutput), "Expected license status output on stderr");
    }

    [SkippableFact]
    public void SetLicense_WithPdfOnly_ShouldOutputLicenseStatus()
    {
        SkipWhenTheHostMustStayUnlicensed();

        var config = ServerConfig.LoadFromArgs(["--pdf"]);

        var exception = Record.Exception(() => LicenseManager.SetLicense(config));

        Assert.Null(exception);
        var errorOutput = _consoleError.ToString();
        Assert.False(string.IsNullOrEmpty(errorOutput), "Expected license status output on stderr");
    }

    [SkippableFact]
    public void SetLicense_WithEmailOnly_ShouldOutputLicenseStatus()
    {
        SkipWhenTheHostMustStayUnlicensed();

        var config = ServerConfig.LoadFromArgs(["--email"]);

        var exception = Record.Exception(() => LicenseManager.SetLicense(config));

        Assert.Null(exception);
        var errorOutput = _consoleError.ToString();
        Assert.False(string.IsNullOrEmpty(errorOutput), "Expected license status output on stderr");
    }

    [SkippableFact]
    public void SetLicense_WithBarCodeOnly_ShouldOutputLicenseStatus()
    {
        SkipWhenTheHostMustStayUnlicensed();

        var config = ServerConfig.LoadFromArgs(["--barcode"]);

        var exception = Record.Exception(() => LicenseManager.SetLicense(config));

        Assert.Null(exception);
        var errorOutput = _consoleError.ToString();
        Assert.False(string.IsNullOrEmpty(errorOutput), "Expected license status output on stderr");
    }

    #endregion

    #region Console Output Tests

    [SkippableFact]
    public void SetLicense_ShouldRestoreConsoleOut()
    {
        SkipWhenTheHostMustStayUnlicensed();

        var config = ServerConfig.LoadFromArgs(["--word"]);

        Console.SetOut(_originalConsoleOut);
        Console.SetError(_consoleError);

        LicenseManager.SetLicense(config);

        Assert.Same(_originalConsoleOut, Console.Out);
    }

    [SkippableFact]
    public void SetLicense_ShouldOutputEvaluationModeMessage()
    {
        SkipWhenTheHostMustStayUnlicensed();

        var config = ServerConfig.LoadFromArgs(["--all"]);

        LicenseManager.SetLicense(config);

        var errorOutput = _consoleError.ToString();
        Assert.True(
            errorOutput.Contains("license", StringComparison.OrdinalIgnoreCase) ||
            errorOutput.Contains("evaluation", StringComparison.OrdinalIgnoreCase));
    }

    #endregion
}
