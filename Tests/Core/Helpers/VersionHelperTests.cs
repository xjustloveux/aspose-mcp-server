using AsposeMcpServer.Core.Helpers;

namespace AsposeMcpServer.Tests.Core.Helpers;

/// <summary>
///     Unit tests for VersionHelper class
/// </summary>
public class VersionHelperTests
{
    [Fact]
    public void GetVersion_ShouldReturnVersionString()
    {
        var version = VersionHelper.GetVersion();

        Assert.NotNull(version);
        Assert.NotEmpty(version);
    }

    [Fact]
    public void GetVersion_ShouldReturnValidFormat()
    {
        var version = VersionHelper.GetVersion();

        // Version should be in format "major.minor.patch"
        var parts = version.Split('.');
        Assert.True(parts.Length >= 2, "Version should have at least major.minor");

        // Each part should be numeric
        foreach (var part in parts)
            Assert.True(int.TryParse(part, out _), $"Version part '{part}' should be numeric");
    }

    [Fact]
    public void GetVersion_ShouldBeCached()
    {
        var version1 = VersionHelper.GetVersion();
        var version2 = VersionHelper.GetVersion();

        Assert.Same(version1, version2);
    }

    [Fact]
    public void GetVersion_ShouldNotBeEmpty()
    {
        var version = VersionHelper.GetVersion();

        Assert.False(string.IsNullOrWhiteSpace(version));
    }

    [Fact]
    public void GetVersion_ShouldNotContainBuildNumber()
    {
        var version = VersionHelper.GetVersion();
        var parts = version.Split('.');

        // Should have at most 3 parts (major.minor.patch)
        Assert.True(parts.Length <= 3, "Version should not have more than 3 parts");
    }
}