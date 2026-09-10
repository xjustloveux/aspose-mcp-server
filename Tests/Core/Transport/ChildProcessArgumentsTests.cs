using System.Text;
using AsposeMcpServer.Core.Transport;

namespace AsposeMcpServer.Tests.Core.Transport;

/// <summary>
///     Guards RB-03 and RB-04: WebSocket mode starts a stdio child per connection. The previous
///     implementation kept only a handful of switches and dropped their values, so the child ran
///     without the path allowlist, without the licence path and with a different tool set than the
///     parent, and it joined the tokens into a single unquoted string.
/// </summary>
public class ChildProcessArgumentsTests
{
    [Fact]
    public void AllowedPathWithValue_ShouldBeForwarded()
    {
        string[] args = ["--ws", "--port", "3000", "--allowed-path", @"D:\srv\my docs", "--all"];

        var result = ChildProcessArguments.BuildChildArguments(args);

        Assert.Equal(["--allowed-path", @"D:\srv\my docs", "--all", "--stdio"], result);
    }

    [Fact]
    public void RepeatedValueOptions_ShouldAllSurvive()
    {
        string[] args = ["--allowed-path", "/a", "--allowed-path", "/b", "--ws"];

        var result = ChildProcessArguments.BuildChildArguments(args);

        Assert.Equal(["--allowed-path", "/a", "--allowed-path", "/b", "--stdio"], result);
    }

    [Fact]
    public void LicenceAndSessionOptions_ShouldKeepTheirValues()
    {
        string[] args = ["--license", @"C:\keys\Aspose.Total.lic", "--session-timeout", "600", "--ws"];

        var result = ChildProcessArguments.BuildChildArguments(args);

        Assert.Equal(
            ["--license", @"C:\keys\Aspose.Total.lic", "--session-timeout", "600", "--stdio"],
            result);
    }

    [Fact]
    public void FeatureFlagsOtherThanTools_ShouldBeForwarded()
    {
        string[] args = ["--ocr", "--email", "--barcode", "--max-extract-all-bytes", "1024", "--websocket"];

        var result = ChildProcessArguments.BuildChildArguments(args);

        Assert.Equal(
            ["--ocr", "--email", "--barcode", "--max-extract-all-bytes", "1024", "--stdio"],
            result);
    }

    [Theory]
    [InlineData("--port", "3000")]
    [InlineData("--host", "0.0.0.0")]
    public void TransportValueOptions_ShouldBeDroppedWithTheirValue(string option, string value)
    {
        string[] args = [option, value, "--word"];

        var result = ChildProcessArguments.BuildChildArguments(args);

        Assert.Equal(["--word", "--stdio"], result);
    }

    [Theory]
    [InlineData("--port:3000")]
    [InlineData("--port=3000")]
    [InlineData("--host:localhost")]
    [InlineData("--host=localhost")]
    public void InlineTransportValues_ShouldBeDropped(string arg)
    {
        var result = ChildProcessArguments.BuildChildArguments([arg, "--word"]);

        Assert.Equal(["--word", "--stdio"], result);
    }

    [Fact]
    public void TransportFlags_ShouldNeverReachTheChild()
    {
        var result = ChildProcessArguments.BuildChildArguments(["--ws", "--http", "--stdio", "--pdf"]);

        Assert.Equal(["--pdf", "--stdio"], result);
        Assert.Single(result, a => a == "--stdio");
    }

    /// <summary>
    ///     A value option with nothing usable after it used to consume whatever came next, so the
    ///     child lost an option the parent was running with (R3-C01).
    /// </summary>
    /// <param name="option">The transport option written without a value.</param>
    [Theory]
    [InlineData("--port")]
    [InlineData("--host")]
    public void ValueOptionWithoutAValue_ShouldBeRefused(string option)
    {
        var exception = Assert.Throws<ArgumentException>(() =>
            ChildProcessArguments.BuildChildArguments(["--word", option]));

        Assert.Contains(option, exception.Message);
    }

    /// <param name="option">The transport option whose value is missing.</param>
    /// <param name="following">The option that follows it and must not be treated as its value.</param>
    [Theory]
    [InlineData("--port", "--allowed-path")]
    [InlineData("--host", "--allowed-path")]
    [InlineData("--port", "--license")]
    [InlineData("--host", "--all")]
    public void ValueOptionFollowedByAnotherOption_ShouldBeRefused(string option, string following)
    {
        var exception = Assert.Throws<ArgumentException>(() =>
            ChildProcessArguments.BuildChildArguments([option, following, @"C:\safe"]));

        Assert.Contains(following, exception.Message);
    }

    [Fact]
    public void HostWithANumericValue_ShouldStillConsumeIt()
    {
        // A host can look like anything that is not an option; only the option shape decides.
        var result = ChildProcessArguments.BuildChildArguments(["--host", "10", "--word"]);

        Assert.Equal(["--word", "--stdio"], result);
    }

    [Fact]
    public void PortWithANonNumericValue_ShouldBeRefused()
    {
        Assert.Throws<ArgumentException>(() =>
            ChildProcessArguments.BuildChildArguments(["--port", "3000x", "--word"]));
    }

    [Fact]
    public void EmptyArguments_ShouldStillSelectStdio()
    {
        Assert.Equal(["--stdio"], ChildProcessArguments.BuildChildArguments([]));
    }

    [Fact]
    public void ChildStartInfo_ShouldCarryEveryTokenAsItsOwnArgument()
    {
        // The tokens were only ever checked as a list returned by BuildChildArguments; nothing
        // looked at what the launch boundary did with them. Joining them into ProcessStartInfo's
        // Arguments string re-splits a path on its spaces, so the child would run with a narrower
        // allowlist than the parent while every existing assertion stayed green (R7-T01).
        IReadOnlyList<string> tokens =
            ["--allowed-path", @"D:\srv\my docs", "--license", @"C:\keys\a b.lic", "--stdio"];

        var startInfo = ChildProcessArguments.CreateChildStartInfo("host.exe", tokens, null, null);

        Assert.Equal(tokens, startInfo.ArgumentList);
        Assert.True(string.IsNullOrEmpty(startInfo.Arguments),
            "Tokens must reach the child through ArgumentList; the Arguments string is re-split by "
            + $"the platform, and it currently holds '{startInfo.Arguments}'.");
        Assert.False(startInfo.UseShellExecute);
    }

    [Fact]
    public void ChildStartInfo_ShouldPassSessionIdentityThroughTheEnvironment()
    {
        var withIdentity = ChildProcessArguments.CreateChildStartInfo(
            "host.exe", ["--stdio"], "group-1", "user-1");

        Assert.Equal("group-1", withIdentity.Environment["ASPOSE_SESSION_GROUP_ID"]);
        Assert.Equal("user-1", withIdentity.Environment["ASPOSE_SESSION_USER_ID"]);

        var withoutIdentity = ChildProcessArguments.CreateChildStartInfo(
            "host.exe", ["--stdio"], null, "");

        Assert.DoesNotContain("ASPOSE_SESSION_GROUP_ID", withoutIdentity.Environment.Keys);
        Assert.DoesNotContain("ASPOSE_SESSION_USER_ID", withoutIdentity.Environment.Keys);
    }

    [Fact]
    public void ChildStartInfo_ShouldRedirectEveryStreamTheBridgeReads()
    {
        var startInfo = ChildProcessArguments.CreateChildStartInfo("host.exe", ["--stdio"], null, null);

        Assert.True(startInfo.RedirectStandardInput);
        Assert.True(startInfo.RedirectStandardOutput);
        Assert.True(startInfo.RedirectStandardError);
        Assert.Equal(Encoding.UTF8, startInfo.StandardInputEncoding);
    }

    [Fact]
    public void ResolveHostCommand_ShouldReturnALaunchableExecutable()
    {
        var (executable, prefix) = ChildProcessArguments.ResolveHostCommand();

        Assert.False(string.IsNullOrEmpty(executable));
        if (Path.GetFileNameWithoutExtension(executable).Equals("dotnet", StringComparison.OrdinalIgnoreCase))
            Assert.Single(prefix);
        else
            Assert.Empty(prefix);
    }
}
