using AsposeMcpServer.Core.Session;

namespace AsposeMcpServer.Tests.Core.Session;

/// <summary>
///     Unit tests for StdioSessionIdentityAccessor class
/// </summary>
public class StdioSessionIdentityAccessorTests
{
    /// <summary>
    ///     Helper to safely set and restore environment variables during tests
    /// </summary>
    private static void WithEnvironmentVariables(string? groupId, string? userId, Action action)
    {
        var originalGroupId = Environment.GetEnvironmentVariable("ASPOSE_SESSION_GROUP_ID");
        var originalUserId = Environment.GetEnvironmentVariable("ASPOSE_SESSION_USER_ID");

        try
        {
            Environment.SetEnvironmentVariable("ASPOSE_SESSION_GROUP_ID", groupId);
            Environment.SetEnvironmentVariable("ASPOSE_SESSION_USER_ID", userId);
            action();
        }
        finally
        {
            Environment.SetEnvironmentVariable("ASPOSE_SESSION_GROUP_ID", originalGroupId);
            Environment.SetEnvironmentVariable("ASPOSE_SESSION_USER_ID", originalUserId);
        }
    }

    #region UserId Environment Variable Tests

    [Fact]
    public void GetCurrentIdentity_WithUserIdOnly_ShouldReturnIdentityWithUserId()
    {
        WithEnvironmentVariables(null, "user1", () =>
        {
            var accessor = new StdioSessionIdentityAccessor();
            var identity = accessor.GetCurrentIdentity();

            Assert.False(identity.IsAnonymous);
            Assert.Null(identity.GroupId);
            Assert.Equal("user1", identity.UserId);
        });
    }

    #endregion

    #region Caching Behavior Tests

    [Fact]
    public void GetCurrentIdentity_CalledMultipleTimes_ShouldReturnSameInstance()
    {
        WithEnvironmentVariables("group1", "user1", () =>
        {
            var accessor = new StdioSessionIdentityAccessor();
            var identity1 = accessor.GetCurrentIdentity();
            var identity2 = accessor.GetCurrentIdentity();

            Assert.Same(identity1, identity2);
        });
    }

    #endregion

    #region Anonymous Identity Tests

    [Fact]
    public void GetCurrentIdentity_NoEnvironmentVariables_ShouldReturnAnonymous()
    {
        WithEnvironmentVariables(null, null, () =>
        {
            var accessor = new StdioSessionIdentityAccessor();
            var identity = accessor.GetCurrentIdentity();

            Assert.True(identity.IsAnonymous);
        });
    }

    [Fact]
    public void GetCurrentIdentity_EmptyEnvironmentVariables_ShouldReturnAnonymous()
    {
        WithEnvironmentVariables("", "", () =>
        {
            var accessor = new StdioSessionIdentityAccessor();
            var identity = accessor.GetCurrentIdentity();

            Assert.True(identity.IsAnonymous);
        });
    }

    #endregion

    #region GroupId Environment Variable Tests

    [Fact]
    public void GetCurrentIdentity_WithGroupIdOnly_ShouldReturnIdentityWithGroupId()
    {
        WithEnvironmentVariables("group1", null, () =>
        {
            var accessor = new StdioSessionIdentityAccessor();
            var identity = accessor.GetCurrentIdentity();

            Assert.False(identity.IsAnonymous);
            Assert.Equal("group1", identity.GroupId);
            Assert.Null(identity.UserId);
        });
    }

    [Fact]
    public void GetCurrentIdentity_WithGroupIdEmpty_ShouldReturnAnonymous()
    {
        WithEnvironmentVariables("", "user1", () =>
        {
            var accessor = new StdioSessionIdentityAccessor();
            var identity = accessor.GetCurrentIdentity();

            Assert.False(identity.IsAnonymous);
            Assert.Null(identity.GroupId);
            Assert.Equal("user1", identity.UserId);
        });
    }

    #endregion

    #region Both Variables Tests

    [Fact]
    public void GetCurrentIdentity_WithBothVariables_ShouldReturnFullIdentity()
    {
        WithEnvironmentVariables("group1", "user1", () =>
        {
            var accessor = new StdioSessionIdentityAccessor();
            var identity = accessor.GetCurrentIdentity();

            Assert.False(identity.IsAnonymous);
            Assert.Equal("group1", identity.GroupId);
            Assert.Equal("user1", identity.UserId);
        });
    }

    [Fact]
    public void GetCurrentIdentity_WithSpecialCharacters_ShouldPreserveValues()
    {
        WithEnvironmentVariables("group:with:colons", "user@domain.com", () =>
        {
            var accessor = new StdioSessionIdentityAccessor();
            var identity = accessor.GetCurrentIdentity();

            Assert.Equal("group:with:colons", identity.GroupId);
            Assert.Equal("user@domain.com", identity.UserId);
        });
    }

    #endregion
}
