using Aspose.Words;
using AsposeMcpServer.Core.Session;

namespace AsposeMcpServer.Tests.Core.Session;

/// <summary>
///     Tests for DocumentSessionManager session isolation functionality
/// </summary>
public class DocumentSessionManagerIsolationTests : IDisposable
{
    private readonly string _testDir;
    private readonly string _testFilePath;

    public DocumentSessionManagerIsolationTests()
    {
        _testDir = Path.Combine(Path.GetTempPath(), $"session_isolation_test_{Guid.NewGuid():N}");
        Directory.CreateDirectory(_testDir);

        _testFilePath = Path.Combine(_testDir, "test.docx");

        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.Write("Test content");
        doc.Save(_testFilePath);
    }

    public void Dispose()
    {
        try
        {
            if (Directory.Exists(_testDir))
                Directory.Delete(_testDir, true);
        }
        catch
        {
            // Ignore cleanup errors
        }
    }

    private DocumentSessionManager CreateManager(SessionIsolationMode mode)
    {
        var config = new SessionConfig
        {
            Enabled = true,
            MaxSessions = 10,
            IsolationMode = mode
        };
        return new DocumentSessionManager(config);
    }

    #region Close Session Authorization Tests

    [Fact]
    public void CloseDocument_UnauthorizedUserCannotClose()
    {
        using var manager = CreateManager(SessionIsolationMode.User);

        var owner = new SessionIdentity { TenantId = "tenant1", UserId = "user1" };
        var other = new SessionIdentity { TenantId = "tenant1", UserId = "user2" };

        var sessionId = manager.OpenDocument(_testFilePath, owner);

        // Other user should not be able to close
        Assert.Throws<KeyNotFoundException>(() => manager.CloseDocument(sessionId, other, true));

        // Owner should be able to close
        manager.CloseDocument(sessionId, owner, true);
    }

    #endregion

    #region Save Document Authorization Tests

    [Fact]
    public void SaveDocument_UnauthorizedUserCannotSave()
    {
        using var manager = CreateManager(SessionIsolationMode.User);

        var owner = new SessionIdentity { TenantId = "tenant1", UserId = "user1" };
        var other = new SessionIdentity { TenantId = "tenant1", UserId = "user2" };

        var sessionId = manager.OpenDocument(_testFilePath, owner);

        // Other user should not be able to save
        Assert.Throws<KeyNotFoundException>(() => manager.SaveDocument(sessionId, other));

        manager.CloseDocument(sessionId, owner, true);
    }

    #endregion

    #region Session Limit Per Owner Tests

    [Fact]
    public void SessionLimitIsPerOwnerInUserMode()
    {
        var config = new SessionConfig
        {
            Enabled = true,
            MaxSessions = 2,
            IsolationMode = SessionIsolationMode.User
        };
        using var manager = new DocumentSessionManager(config);

        var user1 = new SessionIdentity { TenantId = "tenant1", UserId = "user1" };
        var user2 = new SessionIdentity { TenantId = "tenant1", UserId = "user2" };

        // User1 can open 2 sessions
        manager.OpenDocument(_testFilePath, user1);
        manager.OpenDocument(_testFilePath, user1);

        // User1 cannot open a 3rd
        Assert.Throws<InvalidOperationException>(() => manager.OpenDocument(_testFilePath, user1));

        // But user2 can still open sessions
        manager.OpenDocument(_testFilePath, user2);
        manager.OpenDocument(_testFilePath, user2);
    }

    #endregion

    #region None Mode Tests

    [Fact]
    public void NoneMode_AnyoneCanAccessAnySession()
    {
        using var manager = CreateManager(SessionIsolationMode.None);

        var tenant1User1 = new SessionIdentity { TenantId = "tenant1", UserId = "user1" };
        var tenant2User2 = new SessionIdentity { TenantId = "tenant2", UserId = "user2" };

        var sessionId = manager.OpenDocument(_testFilePath, tenant1User1);

        // tenant2/user2 should be able to access it
        var session = manager.TryGetSession(sessionId, tenant2User2);
        Assert.NotNull(session);

        manager.CloseDocument(sessionId, tenant2User2, true);
    }

    [Fact]
    public void NoneMode_ListSessionsShowsAll()
    {
        using var manager = CreateManager(SessionIsolationMode.None);

        var tenant1 = new SessionIdentity { TenantId = "tenant1", UserId = "user1" };
        var tenant2 = new SessionIdentity { TenantId = "tenant2", UserId = "user2" };

        manager.OpenDocument(_testFilePath, tenant1);
        manager.OpenDocument(_testFilePath, tenant2);

        var sessions = manager.ListSessions(tenant1).ToList();
        Assert.Equal(2, sessions.Count);
    }

    #endregion

    #region Tenant Mode Tests

    [Fact]
    public void TenantMode_SameTenantDifferentUserCanAccess()
    {
        using var manager = CreateManager(SessionIsolationMode.Tenant);

        var user1 = new SessionIdentity { TenantId = "tenant1", UserId = "user1" };
        var user2 = new SessionIdentity { TenantId = "tenant1", UserId = "user2" };

        var sessionId = manager.OpenDocument(_testFilePath, user1);

        // Same tenant, different user should access
        var session = manager.TryGetSession(sessionId, user2);
        Assert.NotNull(session);

        manager.CloseDocument(sessionId, user2, true);
    }

    [Fact]
    public void TenantMode_DifferentTenantCannotAccess()
    {
        using var manager = CreateManager(SessionIsolationMode.Tenant);

        var tenant1 = new SessionIdentity { TenantId = "tenant1", UserId = "user1" };
        var tenant2 = new SessionIdentity { TenantId = "tenant2", UserId = "user1" };

        var sessionId = manager.OpenDocument(_testFilePath, tenant1);

        // Different tenant should NOT access
        var session = manager.TryGetSession(sessionId, tenant2);
        Assert.Null(session);

        manager.CloseDocument(sessionId, tenant1, true);
    }

    [Fact]
    public void TenantMode_ListSessionsShowsSameTenantOnly()
    {
        using var manager = CreateManager(SessionIsolationMode.Tenant);

        var tenant1User1 = new SessionIdentity { TenantId = "tenant1", UserId = "user1" };
        var tenant1User2 = new SessionIdentity { TenantId = "tenant1", UserId = "user2" };
        var tenant2 = new SessionIdentity { TenantId = "tenant2", UserId = "user1" };

        manager.OpenDocument(_testFilePath, tenant1User1);
        manager.OpenDocument(_testFilePath, tenant1User2);
        manager.OpenDocument(_testFilePath, tenant2);

        // Tenant1 users should see 2 sessions
        var tenant1Sessions = manager.ListSessions(tenant1User1).ToList();
        Assert.Equal(2, tenant1Sessions.Count);

        // Tenant2 should see 1 session
        var tenant2Sessions = manager.ListSessions(tenant2).ToList();
        Assert.Single(tenant2Sessions);
    }

    #endregion

    #region User Mode Tests

    [Fact]
    public void UserMode_SameUserCanAccess()
    {
        using var manager = CreateManager(SessionIsolationMode.User);

        var user = new SessionIdentity { TenantId = "tenant1", UserId = "user1" };

        var sessionId = manager.OpenDocument(_testFilePath, user);

        // Same user should access
        var session = manager.TryGetSession(sessionId, user);
        Assert.NotNull(session);

        manager.CloseDocument(sessionId, user, true);
    }

    [Fact]
    public void UserMode_SameTenantDifferentUserCannotAccess()
    {
        using var manager = CreateManager(SessionIsolationMode.User);

        var user1 = new SessionIdentity { TenantId = "tenant1", UserId = "user1" };
        var user2 = new SessionIdentity { TenantId = "tenant1", UserId = "user2" };

        var sessionId = manager.OpenDocument(_testFilePath, user1);

        // Same tenant, different user should NOT access
        var session = manager.TryGetSession(sessionId, user2);
        Assert.Null(session);

        manager.CloseDocument(sessionId, user1, true);
    }

    [Fact]
    public void UserMode_DifferentTenantCannotAccess()
    {
        using var manager = CreateManager(SessionIsolationMode.User);

        var tenant1 = new SessionIdentity { TenantId = "tenant1", UserId = "user1" };
        var tenant2 = new SessionIdentity { TenantId = "tenant2", UserId = "user1" };

        var sessionId = manager.OpenDocument(_testFilePath, tenant1);

        // Different tenant should NOT access
        var session = manager.TryGetSession(sessionId, tenant2);
        Assert.Null(session);

        manager.CloseDocument(sessionId, tenant1, true);
    }

    [Fact]
    public void UserMode_ListSessionsShowsOwnOnly()
    {
        using var manager = CreateManager(SessionIsolationMode.User);

        var user1 = new SessionIdentity { TenantId = "tenant1", UserId = "user1" };
        var user2 = new SessionIdentity { TenantId = "tenant1", UserId = "user2" };
        var user3 = new SessionIdentity { TenantId = "tenant2", UserId = "user1" };

        manager.OpenDocument(_testFilePath, user1);
        manager.OpenDocument(_testFilePath, user1);
        manager.OpenDocument(_testFilePath, user2);
        manager.OpenDocument(_testFilePath, user3);

        // User1 should see 2 sessions
        var user1Sessions = manager.ListSessions(user1).ToList();
        Assert.Equal(2, user1Sessions.Count);

        // User2 should see 1 session
        var user2Sessions = manager.ListSessions(user2).ToList();
        Assert.Single(user2Sessions);

        // User3 should see 1 session
        var user3Sessions = manager.ListSessions(user3).ToList();
        Assert.Single(user3Sessions);
    }

    #endregion

    #region Anonymous Access Tests

    [Fact]
    public void AnonymousRequestor_CanAccessAnySession()
    {
        using var manager = CreateManager(SessionIsolationMode.User);

        var owner = new SessionIdentity { TenantId = "tenant1", UserId = "user1" };
        var anonymous = SessionIdentity.GetAnonymous();

        var sessionId = manager.OpenDocument(_testFilePath, owner);

        // Anonymous should access any session (backward compatibility)
        var session = manager.TryGetSession(sessionId, anonymous);
        Assert.NotNull(session);

        manager.CloseDocument(sessionId, anonymous, true);
    }

    [Fact]
    public void AnonymousOwner_CanBeAccessedByAnyone()
    {
        using var manager = CreateManager(SessionIsolationMode.User);

        var anonymous = SessionIdentity.GetAnonymous();
        var user = new SessionIdentity { TenantId = "tenant1", UserId = "user1" };

        var sessionId = manager.OpenDocument(_testFilePath, anonymous);

        // Any user should access anonymous-owned session
        var session = manager.TryGetSession(sessionId, user);
        Assert.NotNull(session);

        manager.CloseDocument(sessionId, user, true);
    }

    [Fact]
    public void AnonymousRequestor_ListSessionsShowsAll()
    {
        using var manager = CreateManager(SessionIsolationMode.User);

        var user1 = new SessionIdentity { TenantId = "tenant1", UserId = "user1" };
        var user2 = new SessionIdentity { TenantId = "tenant2", UserId = "user2" };
        var anonymous = SessionIdentity.GetAnonymous();

        manager.OpenDocument(_testFilePath, user1);
        manager.OpenDocument(_testFilePath, user2);

        // Anonymous should see all sessions
        var sessions = manager.ListSessions(anonymous).ToList();
        Assert.Equal(2, sessions.Count);
    }

    #endregion
}