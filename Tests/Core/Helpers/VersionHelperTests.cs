using AsposeMcpServer.Core.Helpers;

namespace AsposeMcpServer.Tests.Core.Helpers;

/// <summary>
///     Unit tests for VersionHelper class
/// </summary>
public class VersionHelperTests
{
    #region GetVersion Tests

    [Fact]
    public void GetVersion_ShouldReturnNonEmptyVersionString()
    {
        var version = VersionHelper.GetVersion();

        Assert.NotNull(version);
        Assert.NotEmpty(version);
        Assert.False(string.IsNullOrWhiteSpace(version));
    }

    [Fact]
    public void GetVersion_ShouldReturnValidSemanticVersionFormat()
    {
        var version = VersionHelper.GetVersion();
        var parts = version.Split('.');

        Assert.True(parts.Length >= 2, "Version should have at least major.minor");
        Assert.True(parts.Length <= 3, "Version should not have more than 3 parts");

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

    #endregion
}
