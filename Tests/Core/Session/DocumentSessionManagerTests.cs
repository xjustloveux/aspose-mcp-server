using Aspose.Words;
using AsposeMcpServer.Core.Session;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Core.Session;

/// <summary>
///     Unit tests for DocumentSessionManager class
/// </summary>
public class DocumentSessionManagerTests : WordTestBase
{
    private readonly SessionConfig _config = new()
    {
        Enabled = true,
        MaxSessions = 5,
        MaxFileSizeMb = 10,
        IdleTimeoutMinutes = 0
    };

    #region Config Tests

    [Fact]
    public void Config_ShouldBeAccessible()
    {
        using var manager = new DocumentSessionManager(_config);

        Assert.Same(_config, manager.Config);
    }

    #endregion

    #region OpenDocument Tests

    [Fact]
    public void OpenDocument_ShouldReturnSessionId()
    {
        using var manager = new DocumentSessionManager(_config);
        var docPath = CreateWordDocument("test_open.docx");

        var sessionId = manager.OpenDocument(docPath);

        Assert.NotNull(sessionId);
        Assert.StartsWith("sess_", sessionId);
    }

    [Fact]
    public void OpenDocument_WithReadWriteMode_ShouldOpen()
    {
        using var manager = new DocumentSessionManager(_config);
        var docPath = CreateWordDocument("test_rw.docx");

        var sessionId = manager.OpenDocument(docPath);

        Assert.NotNull(sessionId);
        var status = manager.GetSessionStatus(sessionId);
        Assert.Equal("readwrite", status?.Mode);
    }

    [Fact]
    public void OpenDocument_WithReadOnlyMode_ShouldOpen()
    {
        using var manager = new DocumentSessionManager(_config);
        var docPath = CreateWordDocument("test_ro.docx");

        var sessionId = manager.OpenDocument(docPath, "readonly");

        Assert.NotNull(sessionId);
        var status = manager.GetSessionStatus(sessionId);
        Assert.Equal("readonly", status?.Mode);
    }

    [Fact]
    public void OpenDocument_FileNotFound_ShouldThrow()
    {
        using var manager = new DocumentSessionManager(_config);

        Assert.Throws<FileNotFoundException>(() =>
            manager.OpenDocument("nonexistent.docx"));
    }

    [Fact]
    public void OpenDocument_MaxSessionsReached_ShouldThrow()
    {
        var limitedConfig = new SessionConfig { MaxSessions = 2, IdleTimeoutMinutes = 0 };
        using var manager = new DocumentSessionManager(limitedConfig);

        var doc1 = CreateWordDocument("test1.docx");
        var doc2 = CreateWordDocument("test2.docx");
        var doc3 = CreateWordDocument("test3.docx");

        manager.OpenDocument(doc1);
        manager.OpenDocument(doc2);

        var ex = Assert.Throws<InvalidOperationException>(() =>
            manager.OpenDocument(doc3));
        Assert.Contains("Maximum session limit", ex.Message);
    }

    #endregion

    #region GetDocument and GetSession Tests

    [Fact]
    public void GetDocument_ShouldReturnDocument()
    {
        using var manager = new DocumentSessionManager(_config);
        var docPath = CreateWordDocument("test_get.docx");
        var sessionId = manager.OpenDocument(docPath);

        var doc = manager.GetDocument<Document>(sessionId);

        Assert.NotNull(doc);
    }

    [Fact]
    public void GetDocument_InvalidSessionId_ShouldThrow()
    {
        using var manager = new DocumentSessionManager(_config);

        Assert.Throws<KeyNotFoundException>(() =>
            manager.GetDocument<Document>("invalid_session"));
    }

    [Fact]
    public void GetSession_ShouldReturnSession()
    {
        using var manager = new DocumentSessionManager(_config);
        var docPath = CreateWordDocument("test_session.docx");
        var sessionId = manager.OpenDocument(docPath);

        var session = manager.GetSession(sessionId);

        Assert.NotNull(session);
        Assert.Equal(sessionId, session.SessionId);
    }

    [Fact]
    public void GetSession_InvalidSessionId_ShouldThrow()
    {
        using var manager = new DocumentSessionManager(_config);

        Assert.Throws<KeyNotFoundException>(() =>
            manager.GetSession("invalid_session"));
    }

    #endregion

    #region MarkDirty Tests

    [Fact]
    public void MarkDirty_ShouldSetDirtyFlag()
    {
        using var manager = new DocumentSessionManager(_config);
        var docPath = CreateWordDocument("test_dirty.docx");
        var sessionId = manager.OpenDocument(docPath);

        manager.MarkDirty(sessionId);

        var status = manager.GetSessionStatus(sessionId);
        Assert.True(status?.IsDirty);
    }

    [Fact]
    public void MarkDirty_InvalidSessionId_ShouldNotThrow()
    {
        var manager = new DocumentSessionManager(_config);
        try
        {
            var ex = Record.Exception(() => manager.MarkDirty("invalid_session"));
            Assert.Null(ex);
        }
        finally
        {
            manager.Dispose();
        }
    }

    #endregion

    #region SaveDocument Tests

    [Fact]
    public void SaveDocument_ShouldSaveToOriginalPath()
    {
        using var manager = new DocumentSessionManager(_config);
        var docPath = CreateWordDocument("test_save.docx");
        var sessionId = manager.OpenDocument(docPath);

        var doc = manager.GetDocument<Document>(sessionId);
        var builder = new DocumentBuilder(doc);
        builder.Write("Modified content");
        manager.MarkDirty(sessionId);

        manager.SaveDocument(sessionId);

        var savedDoc = new Document(docPath);
        Assert.Contains("Modified content", savedDoc.GetText());
    }

    [Fact]
    public void SaveDocument_WithOutputPath_ShouldSaveToNewPath()
    {
        using var manager = new DocumentSessionManager(_config);
        var docPath = CreateWordDocument("test_save_new.docx");
        var sessionId = manager.OpenDocument(docPath);
        var newPath = Path.Combine(Path.GetDirectoryName(docPath)!, "saved_new.docx");

        manager.SaveDocument(sessionId, newPath);

        Assert.True(File.Exists(newPath));
    }

    [Fact]
    public void SaveDocument_InvalidSessionId_ShouldThrow()
    {
        using var manager = new DocumentSessionManager(_config);

        Assert.Throws<KeyNotFoundException>(() =>
            manager.SaveDocument("invalid_session"));
    }

    [Fact]
    public void SaveDocument_ReadOnlyMode_ShouldThrow()
    {
        using var manager = new DocumentSessionManager(_config);
        var docPath = CreateWordDocument("test_readonly.docx");
        var sessionId = manager.OpenDocument(docPath, "readonly");

        var ex = Assert.Throws<InvalidOperationException>(() =>
            manager.SaveDocument(sessionId));
        Assert.Contains("readonly", ex.Message);
    }

    #endregion

    #region CloseDocument Tests

    [Fact]
    public void CloseDocument_ShouldRemoveSession()
    {
        using var manager = new DocumentSessionManager(_config);
        var docPath = CreateWordDocument("test_close.docx");
        var sessionId = manager.OpenDocument(docPath);

        manager.CloseDocument(sessionId);

        Assert.Throws<KeyNotFoundException>(() =>
            manager.GetSession(sessionId));
    }

    [Fact]
    public void CloseDocument_WithDirty_ShouldSave()
    {
        using var manager = new DocumentSessionManager(_config);
        var docPath = CreateWordDocument("test_close_save.docx");
        var sessionId = manager.OpenDocument(docPath);

        var doc = manager.GetDocument<Document>(sessionId);
        var builder = new DocumentBuilder(doc);
        builder.Write("Modified in session");
        manager.MarkDirty(sessionId);

        manager.CloseDocument(sessionId);

        var savedDoc = new Document(docPath);
        Assert.Contains("Modified in session", savedDoc.GetText());
    }

    [Fact]
    public void CloseDocument_WithDiscardTrue_ShouldNotSave()
    {
        using var manager = new DocumentSessionManager(_config);
        var docPath = CreateWordDocument("test_discard.docx");
        var sessionId = manager.OpenDocument(docPath);

        var doc = manager.GetDocument<Document>(sessionId);
        var builder = new DocumentBuilder(doc);
        builder.Write("Should not be saved");
        manager.MarkDirty(sessionId);

        manager.CloseDocument(sessionId, true);

        var savedDoc = new Document(docPath);
        Assert.DoesNotContain("Should not be saved", savedDoc.GetText());
    }

    [Fact]
    public void CloseDocument_InvalidSessionId_ShouldThrow()
    {
        using var manager = new DocumentSessionManager(_config);

        Assert.Throws<KeyNotFoundException>(() =>
            manager.CloseDocument("invalid_session"));
    }

    #endregion

    #region ListSessions Tests

    [Fact]
    public void ListSessions_ShouldReturnAllSessions()
    {
        using var manager = new DocumentSessionManager(_config);
        var doc1 = CreateWordDocument("list1.docx");
        var doc2 = CreateWordDocument("list2.docx");

        manager.OpenDocument(doc1);
        manager.OpenDocument(doc2);

        var sessions = manager.ListSessions().ToList();

        Assert.Equal(2, sessions.Count);
    }

    [Fact]
    public void ListSessions_Empty_ShouldReturnEmptyList()
    {
        using var manager = new DocumentSessionManager(_config);

        var sessions = manager.ListSessions().ToList();

        Assert.Empty(sessions);
    }

    #endregion

    #region GetSessionStatus Tests

    [Fact]
    public void GetSessionStatus_ShouldReturnSessionInfo()
    {
        using var manager = new DocumentSessionManager(_config);
        var docPath = CreateWordDocument("status.docx");
        var sessionId = manager.OpenDocument(docPath);

        var status = manager.GetSessionStatus(sessionId);

        Assert.NotNull(status);
        Assert.Equal(sessionId, status.SessionId);
        Assert.Equal("word", status.DocumentType);
        Assert.Equal(docPath, status.Path);
        Assert.Equal("readwrite", status.Mode);
        Assert.False(status.IsDirty);
    }

    [Fact]
    public void GetSessionStatus_InvalidSessionId_ShouldReturnNull()
    {
        using var manager = new DocumentSessionManager(_config);

        var status = manager.GetSessionStatus("invalid_session");

        Assert.Null(status);
    }

    #endregion

    #region Memory and Disposal Tests

    [Fact]
    public void GetTotalMemoryMb_ShouldReturnSum()
    {
        using var manager = new DocumentSessionManager(_config);
        var doc1 = CreateWordDocument("mem1.docx");
        var doc2 = CreateWordDocument("mem2.docx");

        manager.OpenDocument(doc1);
        manager.OpenDocument(doc2);

        var totalMemory = manager.GetTotalMemoryMb();

        Assert.True(totalMemory >= 0);
    }

    [Fact]
    public void GetTotalMemoryMb_NoSessions_ShouldReturnZero()
    {
        using var manager = new DocumentSessionManager(_config);

        var totalMemory = manager.GetTotalMemoryMb();

        Assert.Equal(0, totalMemory);
    }

    [Fact]
    public void Dispose_ShouldCleanupAllSessions()
    {
        var manager = new DocumentSessionManager(_config);
        var docPath = CreateWordDocument("dispose.docx");
        manager.OpenDocument(docPath);

        manager.Dispose();

        var sessions = manager.ListSessions().ToList();
        Assert.Empty(sessions);
    }

    #endregion

    #region OnClientDisconnect Tests

    [Fact]
    public void OnClientDisconnect_WithNullClientId_ShouldNotThrow()
    {
        var manager = new DocumentSessionManager(_config);
        try
        {
            var ex = Record.Exception(() => manager.OnClientDisconnect(null));
            Assert.Null(ex);
        }
        finally
        {
            manager.Dispose();
        }
    }

    [Fact]
    public void OnClientDisconnect_WithEmptyClientId_ShouldNotThrow()
    {
        var manager = new DocumentSessionManager(_config);
        try
        {
            var ex = Record.Exception(() => manager.OnClientDisconnect(""));
            Assert.Null(ex);
        }
        finally
        {
            manager.Dispose();
        }
    }

    #endregion
}
