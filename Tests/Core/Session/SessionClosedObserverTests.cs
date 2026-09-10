using Aspose.Words;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Core.Session;

/// <summary>
///     R8-S01: closing a session must not depend on its observers succeeding.
///     <para>
///         <c>SessionClosed</c> was raised synchronously before the session was sealed, and on the
///         explicit close path the session had already been removed from the registry by then. A
///         subscriber that threw therefore skipped the seal, the save and the dispose, leaving a
///         document nothing could reach and nothing would ever release. The event is now raised
///         after the session has actually been dealt with, and each subscriber is isolated.
///     </para>
/// </summary>
public class SessionClosedObserverTests : TestBase
{
    /// <summary>Writes a small Word document for a session to open.</summary>
    /// <param name="name">Fixture file name.</param>
    /// <returns>The document path.</returns>
    private string ADocument(string name)
    {
        var path = CreateTestFilePath(name);
        var document = new Document();
        new DocumentBuilder(document).Writeln("body");
        document.Save(path);
        return path;
    }

    [Fact]
    public void AThrowingSubscriber_ShouldNotStopAnExplicitCloseFromDisposingTheSession()
    {
        var manager = SessionManager;
        var sessionId = manager.OpenDocument(ADocument("observer_close.docx"));
        var session = manager.GetSession(sessionId);

        manager.SessionClosed += (_, _) => throw new InvalidOperationException("observer failed");

        manager.CloseDocument(sessionId, true);

        Assert.True(session.IsDisposed,
            "The session was removed from the registry before the observer ran, so a failure "
            + "there left a document nothing could reach and nothing would release.");
    }

    [Fact]
    public void AThrowingSubscriber_ShouldNotStopShutdownFromDisposingEverySession()
    {
        var manager = SessionManager;
        var first = manager.GetSession(manager.OpenDocument(ADocument("observer_a.docx")));
        var second = manager.GetSession(manager.OpenDocument(ADocument("observer_b.docx")));

        manager.SessionClosed += (_, _) => throw new InvalidOperationException("observer failed");

        manager.OnServerShutdown();

        Assert.True(first.IsDisposed);
        Assert.True(second.IsDisposed);
    }

    [Fact]
    public void OneThrowingSubscriber_ShouldNotStopTheOthersBeingTold()
    {
        var manager = SessionManager;
        var sessionId = manager.OpenDocument(ADocument("observer_many.docx"));
        var told = new List<string>();

        manager.SessionClosed += (_, _) => throw new InvalidOperationException("observer failed");
        manager.SessionClosed += (id, _) => told.Add(id);

        manager.CloseDocument(sessionId, true);

        Assert.Equal([sessionId], told);
    }

    [Fact]
    public void AWorkingSubscriber_ShouldStillBeToldWhenASessionCloses()
    {
        var manager = SessionManager;
        var sessionId = manager.OpenDocument(ADocument("observer_ok.docx"));
        var told = new List<string>();

        manager.SessionClosed += (id, _) => told.Add(id);

        manager.CloseDocument(sessionId, true);

        Assert.Equal([sessionId], told);
    }
}
