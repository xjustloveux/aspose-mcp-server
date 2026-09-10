using AsposeMcpServer.Helpers;

namespace AsposeMcpServer.Tests.Infrastructure;

/// <summary>
///     Guards RB-40: the path shape rules rejected legitimate names because they searched for
///     <c>..</c> and <c>~</c> as substrings rather than as path segments, and applied a
///     Windows-only length ceiling on every platform. Real traversal, control characters,
///     alternate data streams and reserved device names must still be refused.
/// </summary>
public class PathRuleCrossPlatformTests
{
    [Theory]
    [InlineData("v1..2.docx")]
    [InlineData("report~v2.docx")]
    [InlineData("PROGRA~1.docx")]
    [InlineData("notes..final..docx")]
    [InlineData("a~b/c..d.txt")]
    public void LegitimateNamesContainingDotsOrTilde_ShouldBeAccepted(string path)
    {
        Assert.True(SecurityHelper.IsSafeFilePath(path));
    }

    [Theory]
    [InlineData("../secret.txt")]
    [InlineData("a/../../etc/passwd")]
    [InlineData("sub/..")]
    [InlineData("..")]
    public void RealTraversalSegments_ShouldStillBeRejected(string path)
    {
        Assert.False(SecurityHelper.IsSafeFilePath(path));
    }

    [Theory]
    [InlineData("~")]
    [InlineData("~/secrets.txt")]
    public void LeadingTilde_ShouldStillBeRejected(string path)
    {
        Assert.False(SecurityHelper.IsSafeFilePath(path));
    }

    [Theory]
    [InlineData("file.txt:hidden")]
    [InlineData("CON.txt")]
    [InlineData("nul")]
    [InlineData("trailing. ")]
    [InlineData("double//slash.txt")]
    public void PreviouslyRejectedShapes_ShouldRemainRejected(string path)
    {
        Assert.False(SecurityHelper.IsSafeFilePath(path));
    }

    [Fact]
    public void ControlCharacters_ShouldRemainRejected()
    {
        Assert.False(SecurityHelper.IsSafeFilePath("bad\u0007name.txt"));
    }

    [Fact]
    public void PathLongerThanWindowsLimit_ShouldFollowPlatformCeiling()
    {
        var longPath = new string('a', 300) + ".txt";

        var accepted = SecurityHelper.IsSafeFilePath(longPath);

        Assert.Equal(!OperatingSystem.IsWindows(), accepted);
    }

    [Fact]
    public void EmptyOrWhitespace_ShouldBeRejected()
    {
        Assert.False(SecurityHelper.IsSafeFilePath(string.Empty));
        Assert.False(SecurityHelper.IsSafeFilePath("   "));
    }

    [Fact]
    public void RelativePathWithoutAbsoluteOptIn_ShouldRejectRootedPath()
    {
        var rooted = OperatingSystem.IsWindows() ? @"C:\temp\file.txt" : "/tmp/file.txt";

        Assert.False(SecurityHelper.IsSafeFilePath(rooted));
        Assert.True(SecurityHelper.IsSafeFilePath(rooted, true));
    }

    [Fact]
    public void AllowlistMatching_ShouldUsePlatformCaseRules()
    {
        var root = Path.Combine(Path.GetTempPath(), "PathRules_" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        try
        {
            var upper = Path.Combine(root.ToUpperInvariant(), "file.txt");
            var caseSensitivePlatform = !OperatingSystem.IsWindows() && !OperatingSystem.IsMacOS();

            if (caseSensitivePlatform)
                Assert.Throws<ArgumentException>(() =>
                    SecurityHelper.ValidatePathWithinAllowedBases(upper, [root]));
            else
                SecurityHelper.ValidatePathWithinAllowedBases(upper, [root]);
        }
        finally
        {
            Directory.Delete(root, true);
        }
    }
}
