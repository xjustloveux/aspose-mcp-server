using Aspose.Words;
using AsposeMcpServer.Core.Session;
using AsposeMcpServer.Tests.Infrastructure;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Logging.Abstractions;

namespace AsposeMcpServer.Tests.Core.Session;

/// <summary>
///     Covers TEST-07 for the shutdown path. <see cref="SessionLifetimeService" /> had no tests at
///     all, yet it is the only thing that turns application shutdown into session cleanup: if it
///     stops registering the handler, or lets an exception escape, open sessions are lost without a
///     save and nothing else notices.
/// </summary>
public class SessionLifetimeServiceTests : TestBase
{
    /// <summary>Builds a session configuration writing temp files under the test directory.</summary>
    /// <param name="onDisconnect">Behaviour to apply when a session is closed.</param>
    /// <returns>The configuration.</returns>
    private SessionConfig CreateConfig(DisconnectBehavior onDisconnect)
    {
        return new SessionConfig
        {
            Enabled = true,
            OnDisconnect = onDisconnect,
            IdleTimeoutMinutes = 0,
            TempDirectory = TestDir,
            IsolationMode = SessionIsolationMode.None
        };
    }

    /// <summary>Creates a Word document on disk for a session to open.</summary>
    /// <param name="fileName">File name under the test directory.</param>
    /// <returns>The path written.</returns>
    private string CreateWordFile(string fileName)
    {
        var path = CreateTestFilePath(fileName);
        var document = new Document();
        var builder = new DocumentBuilder(document);
        builder.Writeln("original");
        document.Save(path, SaveFormat.Docx);
        return path;
    }

    [Fact]
    public async Task StartAsync_ShouldRegisterTheStoppingHandler()
    {
        using var lifetime = new ControllableLifetime();
        using var manager = new DocumentSessionManager(CreateConfig(DisconnectBehavior.Discard));
        var service = new SessionLifetimeService(manager, lifetime,
            NullLogger<SessionLifetimeService>.Instance);
        var path = CreateWordFile("registered.docx");
        var sessionId = manager.OpenDocument(path);

        await service.StartAsync(CancellationToken.None);
        lifetime.StopApplication();

        Assert.Null(manager.TryGetSession(sessionId, SessionIdentity.GetAnonymous()));
    }

    [Fact]
    public async Task ApplicationStopping_WithAutoSave_ShouldWriteTheSessionBackToItsFile()
    {
        using var lifetime = new ControllableLifetime();
        using var manager = new DocumentSessionManager(CreateConfig(DisconnectBehavior.AutoSave));
        var service = new SessionLifetimeService(manager, lifetime,
            NullLogger<SessionLifetimeService>.Instance);
        var path = CreateWordFile("autosave.docx");
        var sessionId = manager.OpenDocument(path);

        var session = manager.GetSession(sessionId);
        session.Execute(document => { new DocumentBuilder((Document)document).Writeln("added in session"); });
        session.IsDirty = true;

        await service.StartAsync(CancellationToken.None);
        lifetime.StopApplication();

        var saved = new Document(path);
        Assert.Contains("added in session", saved.GetText());
    }

    [Fact]
    public void WithoutStart_ApplicationStopping_ShouldLeaveSessionsAlone()
    {
        using var lifetime = new ControllableLifetime();
        using var manager = new DocumentSessionManager(CreateConfig(DisconnectBehavior.Discard));
        _ = new SessionLifetimeService(manager, lifetime, NullLogger<SessionLifetimeService>.Instance);
        var path = CreateWordFile("unstarted.docx");
        var sessionId = manager.OpenDocument(path);

        lifetime.StopApplication();

        // The handler is registered by StartAsync; without it nothing must touch the sessions.
        Assert.NotNull(manager.TryGetSession(sessionId, SessionIdentity.GetAnonymous()));
    }

    [Fact]
    public async Task ApplicationStopping_WhenCleanupThrows_ShouldNotPropagate()
    {
        using var lifetime = new ControllableLifetime();
        var manager = new DocumentSessionManager(CreateConfig(DisconnectBehavior.Discard));
        var service = new SessionLifetimeService(manager, lifetime,
            NullLogger<SessionLifetimeService>.Instance);
        await service.StartAsync(CancellationToken.None);

        // A cleanup failure has to happen *inside* OnServerShutdown for this to be about the
        // service at all. Disposing the manager first ran shutdown there and then, cleared the
        // registry, and left the stopping callback with nothing to do — so the assertion held
        // whether or not the service caught anything (R8-S02). A logger that throws on the line
        // OnServerShutdown starts with puts the failure where it belongs, and the counter proves
        // the failure path was entered rather than avoided.
        var failures = 0;
        var throwingManager = new DocumentSessionManager(
            CreateConfig(DisconnectBehavior.Discard),
            new ThrowingLoggerFactory(() => failures++));
        var throwingService = new SessionLifetimeService(throwingManager, lifetime,
            NullLogger<SessionLifetimeService>.Instance);
        await throwingService.StartAsync(CancellationToken.None);

        var exception = Record.Exception(() => lifetime.StopApplication());

        Assert.True(failures > 0, "The logger that was supposed to fail was never called, so "
                                  + "nothing exercised the service's failure path.");
        Assert.Null(exception);
    }

    [Fact]
    public async Task StopAsync_ShouldCompleteWithoutTouchingSessions()
    {
        using var lifetime = new ControllableLifetime();
        using var manager = new DocumentSessionManager(CreateConfig(DisconnectBehavior.Discard));
        var service = new SessionLifetimeService(manager, lifetime,
            NullLogger<SessionLifetimeService>.Instance);
        var path = CreateWordFile("stopasync.docx");
        var sessionId = manager.OpenDocument(path);

        await service.StartAsync(CancellationToken.None);
        await service.StopAsync(CancellationToken.None);

        Assert.NotNull(manager.TryGetSession(sessionId, SessionIdentity.GetAnonymous()));
    }

    /// <summary>Minimal <see cref="IHostApplicationLifetime" /> whose stopping token this test controls.</summary>
    private sealed class ControllableLifetime : IHostApplicationLifetime, IDisposable
    {
        private readonly CancellationTokenSource _started = new();
        private readonly CancellationTokenSource _stopped = new();
        private readonly CancellationTokenSource _stopping = new();

        public void Dispose()
        {
            _started.Dispose();
            _stopping.Dispose();
            _stopped.Dispose();
        }

        public CancellationToken ApplicationStarted => _started.Token;
        public CancellationToken ApplicationStopping => _stopping.Token;
        public CancellationToken ApplicationStopped => _stopped.Token;

        public void StopApplication()
        {
            _stopping.Cancel();
        }
    }

    /// <summary>Produces loggers that throw, so a failure happens inside the code under test.</summary>
    /// <param name="onCalled">Counts each throwing call.</param>
    private sealed class ThrowingLoggerFactory(Action onCalled) : ILoggerFactory
    {
        /// <inheritdoc />
        public ILogger CreateLogger(string categoryName)
        {
            return new ThrowingLogger(onCalled);
        }

        /// <inheritdoc />
        public void AddProvider(ILoggerProvider provider)
        {
        }

        /// <inheritdoc />
        public void Dispose()
        {
        }
    }

    /// <summary>A logger whose every write throws.</summary>
    /// <param name="onCalled">Counts each throwing call.</param>
    private sealed class ThrowingLogger(Action onCalled) : ILogger
    {
        /// <inheritdoc />
        public IDisposable? BeginScope<TState>(TState state) where TState : notnull
        {
            return null;
        }

        /// <inheritdoc />
        public bool IsEnabled(LogLevel logLevel)
        {
            return true;
        }

        /// <inheritdoc />
        public void Log<TState>(LogLevel logLevel, EventId eventId, TState state, Exception? exception,
            Func<TState, Exception?, string> formatter)
        {
            onCalled();
            throw new InvalidOperationException("the log sink failed during shutdown");
        }
    }
}
