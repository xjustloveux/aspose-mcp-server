using AsposeMcpServer.Core.Session;

namespace AsposeMcpServer.Tests.Core.Session;

public class SessionIdentityTests
{
    #region IsAnonymous Tests

    [Fact]
    public void IsAnonymous_WhenBothNull_ShouldReturnTrue()
    {
        var identity = new SessionIdentity();
        Assert.True(identity.IsAnonymous);
    }

    [Fact]
    public void IsAnonymous_WhenBothEmpty_ShouldReturnTrue()
    {
        var identity = new SessionIdentity { TenantId = "", UserId = "" };
        Assert.True(identity.IsAnonymous);
    }

    [Fact]
    public void IsAnonymous_WhenTenantIdSet_ShouldReturnFalse()
    {
        var identity = new SessionIdentity { TenantId = "tenant1" };
        Assert.False(identity.IsAnonymous);
    }

    [Fact]
    public void IsAnonymous_WhenUserIdSet_ShouldReturnFalse()
    {
        var identity = new SessionIdentity { UserId = "user1" };
        Assert.False(identity.IsAnonymous);
    }

    [Fact]
    public void Anonymous_ShouldReturnAnonymousIdentity()
    {
        var identity = SessionIdentity.GetAnonymous();
        Assert.True(identity.IsAnonymous);
    }

    #endregion

    #region CanAccess - None Mode Tests

    [Fact]
    public void CanAccess_NoneMode_AnyoneCanAccessAnything()
    {
        var requestor = new SessionIdentity { TenantId = "tenant1", UserId = "user1" };
        var owner = new SessionIdentity { TenantId = "tenant2", UserId = "user2" };

        Assert.True(requestor.CanAccess(owner, SessionIsolationMode.None));
    }

    [Fact]
    public void CanAccess_NoneMode_AnonymousCanAccessOwned()
    {
        var requestor = SessionIdentity.GetAnonymous();
        var owner = new SessionIdentity { TenantId = "tenant1", UserId = "user1" };

        Assert.True(requestor.CanAccess(owner, SessionIsolationMode.None));
    }

    #endregion

    #region CanAccess - Tenant Mode Tests

    [Fact]
    public void CanAccess_TenantMode_SameTenantDifferentUser_ShouldAllowAccess()
    {
        var requestor = new SessionIdentity { TenantId = "tenant1", UserId = "user1" };
        var owner = new SessionIdentity { TenantId = "tenant1", UserId = "user2" };

        Assert.True(requestor.CanAccess(owner, SessionIsolationMode.Tenant));
    }

    [Fact]
    public void CanAccess_TenantMode_DifferentTenant_ShouldDenyAccess()
    {
        var requestor = new SessionIdentity { TenantId = "tenant1", UserId = "user1" };
        var owner = new SessionIdentity { TenantId = "tenant2", UserId = "user1" };

        Assert.False(requestor.CanAccess(owner, SessionIsolationMode.Tenant));
    }

    [Fact]
    public void CanAccess_TenantMode_AnonymousRequestor_ShouldAllowAccess()
    {
        var requestor = SessionIdentity.GetAnonymous();
        var owner = new SessionIdentity { TenantId = "tenant1", UserId = "user1" };

        Assert.True(requestor.CanAccess(owner, SessionIsolationMode.Tenant));
    }

    [Fact]
    public void CanAccess_TenantMode_AnonymousOwner_ShouldAllowAccess()
    {
        var requestor = new SessionIdentity { TenantId = "tenant1", UserId = "user1" };
        var owner = SessionIdentity.GetAnonymous();

        Assert.True(requestor.CanAccess(owner, SessionIsolationMode.Tenant));
    }

    #endregion

    #region CanAccess - User Mode Tests

    [Fact]
    public void CanAccess_UserMode_SameTenantSameUser_ShouldAllowAccess()
    {
        var requestor = new SessionIdentity { TenantId = "tenant1", UserId = "user1" };
        var owner = new SessionIdentity { TenantId = "tenant1", UserId = "user1" };

        Assert.True(requestor.CanAccess(owner, SessionIsolationMode.User));
    }

    [Fact]
    public void CanAccess_UserMode_SameTenantDifferentUser_ShouldDenyAccess()
    {
        var requestor = new SessionIdentity { TenantId = "tenant1", UserId = "user1" };
        var owner = new SessionIdentity { TenantId = "tenant1", UserId = "user2" };

        Assert.False(requestor.CanAccess(owner, SessionIsolationMode.User));
    }

    [Fact]
    public void CanAccess_UserMode_DifferentTenant_ShouldDenyAccess()
    {
        var requestor = new SessionIdentity { TenantId = "tenant1", UserId = "user1" };
        var owner = new SessionIdentity { TenantId = "tenant2", UserId = "user1" };

        Assert.False(requestor.CanAccess(owner, SessionIsolationMode.User));
    }

    [Fact]
    public void CanAccess_UserMode_AnonymousRequestor_ShouldAllowAccess()
    {
        var requestor = SessionIdentity.GetAnonymous();
        var owner = new SessionIdentity { TenantId = "tenant1", UserId = "user1" };

        Assert.True(requestor.CanAccess(owner, SessionIsolationMode.User));
    }

    [Fact]
    public void CanAccess_UserMode_AnonymousOwner_ShouldAllowAccess()
    {
        var requestor = new SessionIdentity { TenantId = "tenant1", UserId = "user1" };
        var owner = SessionIdentity.GetAnonymous();

        Assert.True(requestor.CanAccess(owner, SessionIsolationMode.User));
    }

    #endregion

    #region GetStorageKey Tests

    [Fact]
    public void GetStorageKey_NoneMode_ShouldReturnAnonymous()
    {
        var identity = new SessionIdentity { TenantId = "tenant1", UserId = "user1" };
        Assert.Equal("__anonymous__", identity.GetStorageKey(SessionIsolationMode.None));
    }

    [Fact]
    public void GetStorageKey_AnonymousIdentity_ShouldReturnAnonymous()
    {
        var identity = SessionIdentity.GetAnonymous();
        Assert.Equal("__anonymous__", identity.GetStorageKey(SessionIsolationMode.User));
    }

    [Fact]
    public void GetStorageKey_TenantMode_ShouldReturnTenantKey()
    {
        var identity = new SessionIdentity { TenantId = "tenant1", UserId = "user1" };
        // Base64("tenant1") = "dGVuYW50MQ=="
        Assert.Equal("tenant:dGVuYW50MQ==", identity.GetStorageKey(SessionIsolationMode.Tenant));
    }

    [Fact]
    public void GetStorageKey_UserMode_ShouldReturnUserKey()
    {
        var identity = new SessionIdentity { TenantId = "tenant1", UserId = "user1" };
        // Base64("tenant1") = "dGVuYW50MQ==", Base64("user1") = "dXNlcjE="
        Assert.Equal("user:dGVuYW50MQ==:dXNlcjE=", identity.GetStorageKey(SessionIsolationMode.User));
    }

    [Fact]
    public void GetStorageKey_UserModeWithNullTenant_ShouldHandleGracefully()
    {
        var identity = new SessionIdentity { UserId = "user1" };
        // Base64("user1") = "dXNlcjE="
        Assert.Equal("user::dXNlcjE=", identity.GetStorageKey(SessionIsolationMode.User));
    }

    [Fact]
    public void GetStorageKey_SpecialCharacters_ShouldBeEncodedSafely()
    {
        // Test that special characters like ":" are safely encoded
        var identity = new SessionIdentity { TenantId = "tenant:with:colons", UserId = "user:1" };
        var key = identity.GetStorageKey(SessionIsolationMode.User);

        // The key should not contain raw colons from the IDs (only the separator colons)
        var parts = key.Split(':');
        Assert.Equal(3, parts.Length); // "user", encoded_tenant, encoded_user
        Assert.Equal("user", parts[0]);
    }

    #endregion

    #region ToString Tests

    [Fact]
    public void ToString_Anonymous_ShouldReturnAnonymous()
    {
        var identity = SessionIdentity.GetAnonymous();
        Assert.Equal("Anonymous", identity.ToString());
    }

    [Fact]
    public void ToString_WithBothValues_ShouldReturnFormattedString()
    {
        var identity = new SessionIdentity { TenantId = "tenant1", UserId = "user1" };
        Assert.Equal("Tenant=tenant1, User=user1", identity.ToString());
    }

    [Fact]
    public void ToString_WithNullTenant_ShouldShowNull()
    {
        var identity = new SessionIdentity { UserId = "user1" };
        Assert.Equal("Tenant=(null), User=user1", identity.ToString());
    }

    #endregion

    #region Equality Tests

    [Fact]
    public void Equals_SameValues_ShouldReturnTrue()
    {
        var identity1 = new SessionIdentity { TenantId = "tenant1", UserId = "user1" };
        var identity2 = new SessionIdentity { TenantId = "tenant1", UserId = "user1" };

        Assert.True(identity1.Equals(identity2));
    }

    [Fact]
    public void Equals_DifferentValues_ShouldReturnFalse()
    {
        var identity1 = new SessionIdentity { TenantId = "tenant1", UserId = "user1" };
        var identity2 = new SessionIdentity { TenantId = "tenant1", UserId = "user2" };

        Assert.False(identity1.Equals(identity2));
    }

    [Fact]
    public void Equals_BothAnonymous_ShouldReturnTrue()
    {
        var identity1 = SessionIdentity.GetAnonymous();
        var identity2 = new SessionIdentity();

        Assert.True(identity1.Equals(identity2));
    }

    [Fact]
    public void GetHashCode_SameValues_ShouldReturnSameHash()
    {
        var identity1 = new SessionIdentity { TenantId = "tenant1", UserId = "user1" };
        var identity2 = new SessionIdentity { TenantId = "tenant1", UserId = "user1" };

        Assert.Equal(identity1.GetHashCode(), identity2.GetHashCode());
    }

    #endregion
}