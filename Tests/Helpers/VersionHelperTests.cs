using System.Reflection;
using AsposeMcpServer.Helpers;

namespace AsposeMcpServer.Tests.Helpers;

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

    #region Cache Reset Tests

    [Fact]
    public void GetVersion_WhenCacheIsNull_ReinitializesFromAssembly()
    {
        var field = typeof(VersionHelper).GetField("_version", BindingFlags.Static | BindingFlags.NonPublic)!;
        var original = (string?)field.GetValue(null);
        try
        {
            field.SetValue(null, null);

            var result = VersionHelper.GetVersion();

            Assert.NotNull(result);
            Assert.NotEmpty(result);
            var parts = result.Split('.');
            Assert.True(parts.Length is >= 2 and <= 3);
            foreach (var part in parts)
                Assert.True(int.TryParse(part, out _), $"Version part '{part}' should be numeric");
        }
        finally
        {
            field.SetValue(null, original);
        }
    }

    [Fact]
    public void GetVersion_AfterCacheReset_CachesNewValue()
    {
        var field = typeof(VersionHelper).GetField("_version", BindingFlags.Static | BindingFlags.NonPublic)!;
        var original = (string?)field.GetValue(null);
        try
        {
            field.SetValue(null, null);

            var first = VersionHelper.GetVersion();
            var second = VersionHelper.GetVersion();

            Assert.Same(first, second);
            Assert.NotNull(field.GetValue(null));
        }
        finally
        {
            field.SetValue(null, original);
        }
    }

    [Fact]
    public void GetVersion_WhenCacheHasValue_ReturnsCachedDirectly()
    {
        var field = typeof(VersionHelper).GetField("_version", BindingFlags.Static | BindingFlags.NonPublic)!;
        var original = (string?)field.GetValue(null);
        try
        {
            field.SetValue(null, "test.cached");

            var result = VersionHelper.GetVersion();

            Assert.Equal("test.cached", result);
        }
        finally
        {
            field.SetValue(null, original);
        }
    }

    #endregion
}
