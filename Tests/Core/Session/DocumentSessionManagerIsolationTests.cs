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

    private static DocumentSessionManager CreateManager(SessionIsolationMode mode)
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
        using var manager = CreateManager(SessionIsolationMode.Group);

        var owner = new SessionIdentity { GroupId = "group1", UserId = "user1" };
        var other = new SessionIdentity { GroupId = "group2", UserId = "user2" };

        var sessionId = manager.OpenDocument(_testFilePath, owner);

        // Other group should not be able to close
        Assert.Throws<KeyNotFoundException>(() => manager.CloseDocument(sessionId, other, true));

        // Owner should be able to close
        manager.CloseDocument(sessionId, owner, true);
    }

    #endregion

    #region Save Document Authorization Tests

    [Fact]
    public void SaveDocument_UnauthorizedUserCannotSave()
    {
        using var manager = CreateManager(SessionIsolationMode.Group);

        var owner = new SessionIdentity { GroupId = "group1", UserId = "user1" };
        var other = new SessionIdentity { GroupId = "group2", UserId = "user2" };

        var sessionId = manager.OpenDocument(_testFilePath, owner);

        // Other group should not be able to save
        Assert.Throws<KeyNotFoundException>(() => manager.SaveDocument(sessionId, other));

        manager.CloseDocument(sessionId, owner, true);
    }

    #endregion

    #region Session Limit Per Owner Tests

    [Fact]
    public void SessionLimitIsPerOwnerInGroupMode()
    {
        var config = new SessionConfig
        {
            Enabled = true,
            MaxSessions = 2,
            IsolationMode = SessionIsolationMode.Group
        };
        using var manager = new DocumentSessionManager(config);

        var group1 = new SessionIdentity { GroupId = "group1", UserId = "user1" };
        var group2 = new SessionIdentity { GroupId = "group2", UserId = "user2" };

        // Group1 can open 2 sessions
        manager.OpenDocument(_testFilePath, group1);
        manager.OpenDocument(_testFilePath, group1);

        // Group1 cannot open a 3rd
        Assert.Throws<InvalidOperationException>(() => manager.OpenDocument(_testFilePath, group1));

        // But group2 can still open sessions
        manager.OpenDocument(_testFilePath, group2);
        manager.OpenDocument(_testFilePath, group2);
    }

    #endregion

    #region None Mode Tests

    [Fact]
    public void NoneMode_AnyoneCanAccessAnySession()
    {
        using var manager = CreateManager(SessionIsolationMode.None);

        var group1User1 = new SessionIdentity { GroupId = "group1", UserId = "user1" };
        var group2User2 = new SessionIdentity { GroupId = "group2", UserId = "user2" };

        var sessionId = manager.OpenDocument(_testFilePath, group1User1);

        // group2/user2 should be able to access it
        var session = manager.TryGetSession(sessionId, group2User2);
        Assert.NotNull(session);

        manager.CloseDocument(sessionId, group2User2, true);
    }

    [Fact]
    public void NoneMode_ListSessionsShowsAll()
    {
        using var manager = CreateManager(SessionIsolationMode.None);

        var group1 = new SessionIdentity { GroupId = "group1", UserId = "user1" };
        var group2 = new SessionIdentity { GroupId = "group2", UserId = "user2" };

        manager.OpenDocument(_testFilePath, group1);
        manager.OpenDocument(_testFilePath, group2);

        var sessions = manager.ListSessions(group1).ToList();
        Assert.Equal(2, sessions.Count);
    }

    #endregion

    #region Group Mode Tests

    [Fact]
    public void GroupMode_SameGroupDifferentUserCanAccess()
    {
        using var manager = CreateManager(SessionIsolationMode.Group);

        var user1 = new SessionIdentity { GroupId = "group1", UserId = "user1" };
        var user2 = new SessionIdentity { GroupId = "group1", UserId = "user2" };

        var sessionId = manager.OpenDocument(_testFilePath, user1);

        // Same group, different user should access
        var session = manager.TryGetSession(sessionId, user2);
        Assert.NotNull(session);

        manager.CloseDocument(sessionId, user2, true);
    }

    [Fact]
    public void GroupMode_DifferentGroupCannotAccess()
    {
        using var manager = CreateManager(SessionIsolationMode.Group);

        var group1 = new SessionIdentity { GroupId = "group1", UserId = "user1" };
        var group2 = new SessionIdentity { GroupId = "group2", UserId = "user1" };

        var sessionId = manager.OpenDocument(_testFilePath, group1);

        // Different group should NOT access
        var session = manager.TryGetSession(sessionId, group2);
        Assert.Null(session);

        manager.CloseDocument(sessionId, group1, true);
    }

    [Fact]
    public void GroupMode_ListSessionsShowsSameGroupOnly()
    {
        using var manager = CreateManager(SessionIsolationMode.Group);

        var group1User1 = new SessionIdentity { GroupId = "group1", UserId = "user1" };
        var group1User2 = new SessionIdentity { GroupId = "group1", UserId = "user2" };
        var group2 = new SessionIdentity { GroupId = "group2", UserId = "user1" };

        manager.OpenDocument(_testFilePath, group1User1);
        manager.OpenDocument(_testFilePath, group1User2);
        manager.OpenDocument(_testFilePath, group2);

        // Group1 users should see 2 sessions
        var group1Sessions = manager.ListSessions(group1User1).ToList();
        Assert.Equal(2, group1Sessions.Count);

        // Group2 should see 1 session
        var group2Sessions = manager.ListSessions(group2).ToList();
        Assert.Single(group2Sessions);
    }

    [Fact]
    public void GroupMode_SameGroupSameUserCanAccess()
    {
        using var manager = CreateManager(SessionIsolationMode.Group);

        var user = new SessionIdentity { GroupId = "group1", UserId = "user1" };

        var sessionId = manager.OpenDocument(_testFilePath, user);

        // Same user should access
        var session = manager.TryGetSession(sessionId, user);
        Assert.NotNull(session);

        manager.CloseDocument(sessionId, user, true);
    }

    #endregion

    #region Anonymous Access Tests

    [Fact]
    public void AnonymousRequestor_CannotAccessOwnedSession()
    {
        using var manager = CreateManager(SessionIsolationMode.Group);

        var owner = new SessionIdentity { GroupId = "group1", UserId = "user1" };
        var anonymous = SessionIdentity.GetAnonymous();

        var sessionId = manager.OpenDocument(_testFilePath, owner);

        var session = manager.TryGetSession(sessionId, anonymous);
        Assert.Null(session);

        manager.CloseDocument(sessionId, owner, true);
    }

    [Fact]
    public void AnonymousOwner_CannotBeAccessedByOtherGroup()
    {
        using var manager = CreateManager(SessionIsolationMode.Group);

        var anonymous = SessionIdentity.GetAnonymous();
        var user = new SessionIdentity { GroupId = "group1", UserId = "user1" };

        var sessionId = manager.OpenDocument(_testFilePath, anonymous);

        var session = manager.TryGetSession(sessionId, user);
        Assert.Null(session);

        manager.CloseDocument(sessionId, anonymous, true);
    }

    [Fact]
    public void AnonymousRequestor_ListSessionsShowsOnlyAnonymousSessions()
    {
        using var manager = CreateManager(SessionIsolationMode.Group);

        var user1 = new SessionIdentity { GroupId = "group1", UserId = "user1" };
        var user2 = new SessionIdentity { GroupId = "group2", UserId = "user2" };
        var anonymous = SessionIdentity.GetAnonymous();

        manager.OpenDocument(_testFilePath, user1);
        manager.OpenDocument(_testFilePath, user2);
        manager.OpenDocument(_testFilePath, anonymous);

        var sessions = manager.ListSessions(anonymous).ToList();
        Assert.Single(sessions);
    }

    #endregion
}
