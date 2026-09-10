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
        var identity = new SessionIdentity { GroupId = "", UserId = "" };
        Assert.True(identity.IsAnonymous);
    }

    [Fact]
    public void IsAnonymous_WhenGroupIdSet_ShouldReturnFalse()
    {
        var identity = new SessionIdentity { GroupId = "group1" };
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
        var requestor = new SessionIdentity { GroupId = "group1", UserId = "user1" };
        var owner = new SessionIdentity { GroupId = "group2", UserId = "user2" };

        Assert.True(requestor.CanAccess(owner, SessionIsolationMode.None));
    }

    [Fact]
    public void CanAccess_NoneMode_AnonymousCanAccessOwned()
    {
        var requestor = SessionIdentity.GetAnonymous();
        var owner = new SessionIdentity { GroupId = "group1", UserId = "user1" };

        Assert.True(requestor.CanAccess(owner, SessionIsolationMode.None));
    }

    #endregion

    #region CanAccess - Group Mode Tests

    [Fact]
    public void CanAccess_GroupMode_SameGroupDifferentUser_ShouldAllowAccess()
    {
        var requestor = new SessionIdentity { GroupId = "group1", UserId = "user1" };
        var owner = new SessionIdentity { GroupId = "group1", UserId = "user2" };

        Assert.True(requestor.CanAccess(owner, SessionIsolationMode.Group));
    }

    [Fact]
    public void CanAccess_GroupMode_DifferentGroup_ShouldDenyAccess()
    {
        var requestor = new SessionIdentity { GroupId = "group1", UserId = "user1" };
        var owner = new SessionIdentity { GroupId = "group2", UserId = "user1" };

        Assert.False(requestor.CanAccess(owner, SessionIsolationMode.Group));
    }

    [Fact]
    public void CanAccess_GroupMode_AnonymousRequestor_ShouldDenyAccess()
    {
        var requestor = SessionIdentity.GetAnonymous();
        var owner = new SessionIdentity { GroupId = "group1", UserId = "user1" };

        Assert.False(requestor.CanAccess(owner, SessionIsolationMode.Group));
    }

    [Fact]
    public void CanAccess_GroupMode_AnonymousOwner_ShouldDenyAccess()
    {
        var requestor = new SessionIdentity { GroupId = "group1", UserId = "user1" };
        var owner = SessionIdentity.GetAnonymous();

        Assert.False(requestor.CanAccess(owner, SessionIsolationMode.Group));
    }

    [Fact]
    public void CanAccess_GroupMode_BothAnonymous_ShouldAllowAccess()
    {
        var requestor = SessionIdentity.GetAnonymous();
        var owner = SessionIdentity.GetAnonymous();

        Assert.True(requestor.CanAccess(owner, SessionIsolationMode.Group));
    }

    [Fact]
    public void CanAccess_GroupMode_SameGroupSameUser_ShouldAllowAccess()
    {
        var requestor = new SessionIdentity { GroupId = "group1", UserId = "user1" };
        var owner = new SessionIdentity { GroupId = "group1", UserId = "user1" };

        Assert.True(requestor.CanAccess(owner, SessionIsolationMode.Group));
    }

    /// <summary>
    ///     An identity that authenticated but carries no group claim is not anonymous, yet the
    ///     authorisation check compared raw group ids, so two such principals matched on
    ///     <c>null == null</c> and could reach each other's sessions (R2-S03). Several auth modes
    ///     produce exactly that shape: a JWT whose configured group claim is absent, or a gateway
    ///     request without the group header.
    /// </summary>
    [Fact]
    public void CanAccess_GroupMode_DifferentUsersWithNoGroup_ShouldDenyAccess()
    {
        var requestor = new SessionIdentity { UserId = "alice" };
        var owner = new SessionIdentity { UserId = "bob" };

        Assert.False(requestor.CanAccess(owner, SessionIsolationMode.Group));
    }

    [Fact]
    public void CanAccess_GroupMode_SameUserWithNoGroup_ShouldAllowAccess()
    {
        // Falling back to the subject must not lock a principal out of its own sessions.
        var requestor = new SessionIdentity { UserId = "alice" };
        var owner = new SessionIdentity { UserId = "alice" };

        Assert.True(requestor.CanAccess(owner, SessionIsolationMode.Group));
    }

    [Fact]
    public void CanAccess_GroupMode_GrouplessUserAndAnonymous_ShouldDenyAccess()
    {
        var requestor = new SessionIdentity { UserId = "alice" };

        Assert.False(requestor.CanAccess(SessionIdentity.GetAnonymous(), SessionIsolationMode.Group));
        Assert.False(SessionIdentity.GetAnonymous().CanAccess(requestor, SessionIsolationMode.Group));
    }

    #endregion

    #region GetStorageKey Tests

    [Fact]
    public void GetStorageKey_GroupMode_DifferentUsersWithNoGroup_ShouldNotCollide()
    {
        var alice = new SessionIdentity { UserId = "alice" };
        var bob = new SessionIdentity { UserId = "bob" };

        Assert.NotEqual(alice.GetStorageKey(SessionIsolationMode.Group),
            bob.GetStorageKey(SessionIsolationMode.Group));
        Assert.NotEqual("__anonymous__", alice.GetStorageKey(SessionIsolationMode.Group));
    }

    [Fact]
    public void GetStorageKey_NoneMode_ShouldReturnAnonymous()
    {
        var identity = new SessionIdentity { GroupId = "group1", UserId = "user1" };
        Assert.Equal("__anonymous__", identity.GetStorageKey(SessionIsolationMode.None));
    }

    [Fact]
    public void GetStorageKey_AnonymousIdentity_ShouldReturnAnonymous()
    {
        var identity = SessionIdentity.GetAnonymous();
        Assert.Equal("__anonymous__", identity.GetStorageKey(SessionIsolationMode.Group));
    }

    [Fact]
    public void GetStorageKey_GroupMode_ShouldReturnGroupKey()
    {
        var identity = new SessionIdentity { GroupId = "group1", UserId = "user1" };
        // Base64("group1") = "Z3JvdXAx"
        Assert.Equal("group:Z3JvdXAx", identity.GetStorageKey(SessionIsolationMode.Group));
    }

    /// <summary>
    ///     A group-less authenticated identity is stored under its own subject, not under the
    ///     empty group.
    ///     <para>
    ///         This case previously asserted <c>"group:"</c> under the name "ShouldHandleGracefully",
    ///         which made the shared empty-group bucket look deliberate: every principal that
    ///         authenticated without a group claim landed in it together (R2-S03).
    ///     </para>
    /// </summary>
    [Fact]
    public void GetStorageKey_GroupModeWithNullGroup_ShouldUseTheSubject()
    {
        var identity = new SessionIdentity { UserId = "user1" };
        // Base64("user1") = "dXNlcjE="
        Assert.Equal("user:dXNlcjE=", identity.GetStorageKey(SessionIsolationMode.Group));
    }

    [Fact]
    public void GetStorageKey_SpecialCharacters_ShouldBeEncodedSafely()
    {
        var identity = new SessionIdentity { GroupId = "group:with:colons", UserId = "user:1" };
        var key = identity.GetStorageKey(SessionIsolationMode.Group);

        // The key should not contain raw colons from the IDs (only the separator colon)
        var parts = key.Split(':');
        Assert.Equal(2, parts.Length); // "group", encoded_group
        Assert.Equal("group", parts[0]);
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
        var identity = new SessionIdentity { GroupId = "group1", UserId = "user1" };
        Assert.Equal("Group=group1, User=user1", identity.ToString());
    }

    [Fact]
    public void ToString_WithNullGroup_ShouldShowNull()
    {
        var identity = new SessionIdentity { UserId = "user1" };
        Assert.Equal("Group=(null), User=user1", identity.ToString());
    }

    #endregion

    #region Equality Tests

    [Fact]
    public void Equals_SameValues_ShouldReturnTrue()
    {
        var identity1 = new SessionIdentity { GroupId = "group1", UserId = "user1" };
        var identity2 = new SessionIdentity { GroupId = "group1", UserId = "user1" };

        Assert.True(identity1.Equals(identity2));
    }

    [Fact]
    public void Equals_DifferentValues_ShouldReturnFalse()
    {
        var identity1 = new SessionIdentity { GroupId = "group1", UserId = "user1" };
        var identity2 = new SessionIdentity { GroupId = "group1", UserId = "user2" };

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
        var identity1 = new SessionIdentity { GroupId = "group1", UserId = "user1" };
        var identity2 = new SessionIdentity { GroupId = "group1", UserId = "user1" };

        Assert.Equal(identity1.GetHashCode(), identity2.GetHashCode());
    }

    #endregion
}
