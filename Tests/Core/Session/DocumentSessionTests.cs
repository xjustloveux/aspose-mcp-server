using AsposeMcpServer.Core.Session;

namespace AsposeMcpServer.Tests.Core.Session;

/// <summary>
///     Unit tests for DocumentSession class
/// </summary>
public class DocumentSessionTests : IDisposable
{
    private readonly List<DocumentSession> _sessions = new();

    public void Dispose()
    {
        foreach (var session in _sessions) session.Dispose();
    }

    private DocumentSession CreateSession(string sessionId = "sess_test", string path = "test.docx",
        DocumentType type = DocumentType.Word, string mode = "readwrite")
    {
        var mockDocument = new MockDocument();
        var session = new DocumentSession(sessionId, path, type, mockDocument, mode);
        _sessions.Add(session);
        return session;
    }

    [Fact]
    public void Constructor_ShouldSetAllProperties()
    {
        var session = CreateSession("sess_123", "/path/to/doc.docx", DocumentType.Word, "readonly");

        Assert.Equal("sess_123", session.SessionId);
        Assert.Equal("/path/to/doc.docx", session.Path);
        Assert.Equal(DocumentType.Word, session.Type);
        Assert.Equal("readonly", session.Mode);
        Assert.NotNull(session.Document);
    }

    [Fact]
    public void Constructor_ShouldSetTimestamps()
    {
        var beforeCreate = DateTime.UtcNow;
        var session = CreateSession();
        var afterCreate = DateTime.UtcNow;

        Assert.True(session.OpenedAt >= beforeCreate);
        Assert.True(session.OpenedAt <= afterCreate);
        Assert.True(session.LastAccessedAt >= beforeCreate);
        Assert.True(session.LastAccessedAt <= afterCreate);
    }

    [Fact]
    public void IsDirty_DefaultShouldBeFalse()
    {
        var session = CreateSession();

        Assert.False(session.IsDirty);
    }

    [Fact]
    public void IsDirty_ShouldBeSettable()
    {
        var session = CreateSession();

        session.IsDirty = true;

        Assert.True(session.IsDirty);
    }

    [Fact]
    public void ClientId_DefaultShouldBeNull()
    {
        var session = CreateSession();

        Assert.Null(session.ClientId);
    }

    [Fact]
    public void ClientId_ShouldBeSettable()
    {
        var session = CreateSession();

        session.ClientId = "client_123";

        Assert.Equal("client_123", session.ClientId);
    }

    [Fact]
    public void EstimatedMemoryBytes_DefaultShouldBeZero()
    {
        var session = CreateSession();

        Assert.Equal(0, session.EstimatedMemoryBytes);
    }

    [Fact]
    public void EstimatedMemoryBytes_ShouldBeSettable()
    {
        var session = CreateSession();

        session.EstimatedMemoryBytes = 1024 * 1024;

        Assert.Equal(1024 * 1024, session.EstimatedMemoryBytes);
    }

    [Fact]
    public void GetDocument_ShouldReturnTypedDocument()
    {
        var session = CreateSession();

        var doc = session.GetDocument<MockDocument>();

        Assert.NotNull(doc);
        Assert.IsType<MockDocument>(doc);
    }

    [Fact]
    public void GetDocument_ShouldUpdateLastAccessedAt()
    {
        var session = CreateSession();
        var initialTime = session.LastAccessedAt;

        Thread.Sleep(10);
        session.GetDocument<MockDocument>();

        Assert.True(session.LastAccessedAt > initialTime);
    }

    [Fact]
    public void GetDocument_WithWrongType_ShouldThrow()
    {
        var session = CreateSession();

        Assert.Throws<InvalidCastException>(() =>
            session.GetDocument<string>());
    }

    [Fact]
    public async Task ExecuteAsync_ShouldExecuteOperation()
    {
        var session = CreateSession();

        var result = await session.ExecuteAsync(doc =>
        {
            var mockDoc = (MockDocument)doc;
            return mockDoc.Value + 10;
        });

        Assert.Equal(52, result);
    }

    [Fact]
    public async Task ExecuteAsync_ShouldUpdateLastAccessedAt()
    {
        var session = CreateSession();
        var initialTime = session.LastAccessedAt;

        await Task.Delay(10);
        await session.ExecuteAsync(_ => 1);

        Assert.True(session.LastAccessedAt > initialTime);
    }

    [Fact]
    public async Task ExecuteAsyncWithTask_ShouldExecuteAsyncOperation()
    {
        var session = CreateSession();

        var result = await session.ExecuteAsync(async doc =>
        {
            await Task.Delay(1);
            var mockDoc = (MockDocument)doc;
            return mockDoc.Value * 2;
        });

        Assert.Equal(84, result);
    }

    [Fact]
    public void Dispose_ShouldDisposeDocument()
    {
        var mockDocument = new MockDocument();
        var session = new DocumentSession("sess_dispose", "test.docx", DocumentType.Word, mockDocument, "readwrite");

        session.Dispose();

        Assert.True(mockDocument.IsDisposed);
    }

    [Fact]
    public void Dispose_ShouldBeIdempotent()
    {
        var mockDocument = new MockDocument();
        var session = new DocumentSession("sess_dispose2", "test.docx", DocumentType.Word, mockDocument, "readwrite");

        session.Dispose();
        session.Dispose();

        Assert.True(mockDocument.IsDisposed);
        Assert.Equal(1, mockDocument.DisposeCount);
    }

    /// <summary>
    ///     Mock document for testing
    /// </summary>
    private class MockDocument : IDisposable
    {
        public int Value { get; } = 42;
        public bool IsDisposed { get; private set; }
        public int DisposeCount { get; private set; }

        public void Dispose()
        {
            if (!IsDisposed)
            {
                IsDisposed = true;
                DisposeCount++;
            }
        }
    }
}