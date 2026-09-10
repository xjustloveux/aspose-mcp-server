using Aspose.Words;
using AsposeMcpServer.Core.Conversion;
using AsposeMcpServer.Core.Extension;
using AsposeMcpServer.Core.Session;
using AsposeMcpServer.Results.Extension;
using AsposeMcpServer.Tests.Infrastructure;
using Microsoft.Extensions.Logging.Abstractions;

namespace AsposeMcpServer.Tests.Core.Extension;

/// <summary>
///     Covers TEST-07 for <see cref="ExtensionSessionBridge" />, which stood at roughly a tenth of
///     its lines covered. The binding surface is what decides whether a session's document is handed
///     to an out-of-process extension, so its refusals matter as much as its successes: an
///     unknown session, an unknown extension, a closed bridge and another owner's session must all
///     be turned away rather than bound.
/// </summary>
public class ExtensionSessionBridgeTests : TestBase
{
    /// <summary>Builds an extension configuration rooted in the test directory.</summary>
    /// <returns>The configuration.</returns>
    private ExtensionConfig CreateExtensionConfig()
    {
        return new ExtensionConfig
        {
            Enabled = true,
            TempDirectory = TestDir
        };
    }

    /// <summary>Builds a session configuration rooted in the test directory.</summary>
    /// <param name="isolation">Isolation mode governing who may see a session.</param>
    /// <returns>The configuration.</returns>
    private SessionConfig CreateSessionConfig(SessionIsolationMode isolation = SessionIsolationMode.None)
    {
        return new SessionConfig
        {
            Enabled = true,
            IdleTimeoutMinutes = 0,
            TempDirectory = TestDir,
            IsolationMode = isolation
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
        builder.Writeln("content");
        document.Save(path, SaveFormat.Docx);
        return path;
    }

    /// <summary>Assembles a bridge over a fresh session manager and extension manager.</summary>
    /// <param name="sessionConfig">Session configuration to use.</param>
    /// <returns>The bridge plus the managers it was built on, all owned by the caller.</returns>
    private (ExtensionSessionBridge Bridge, DocumentSessionManager Sessions, ExtensionManager Extensions)
        CreateBridge(SessionConfig? sessionConfig = null)
    {
        var extensionConfig = CreateExtensionConfig();
        var sessions = new DocumentSessionManager(sessionConfig ?? CreateSessionConfig());
        var snapshots = new SnapshotManager(extensionConfig, NullLogger<SnapshotManager>.Instance);
        var extensions = new ExtensionManager(extensionConfig, snapshots,
            NullLoggerFactory.Instance, NullLogger<ExtensionManager>.Instance);
        var bridge = new ExtensionSessionBridge(extensionConfig, sessions, extensions,
            NullLogger<ExtensionSessionBridge>.Instance);
        return (bridge, sessions, extensions);
    }

    [Fact]
    public async Task BindAsync_WithUnknownSession_ShouldReportSessionNotFound()
    {
        var (bridge, sessions, extensions) = CreateBridge();
        using var _ = bridge;
        using var __ = sessions;
        await using var ___ = extensions;

        var result = await bridge.BindAsync("no-such-session", "any-extension", "pdf",
            ConversionOptions.WithoutAHost(), SessionIdentity.GetAnonymous());

        Assert.False(result.IsSuccess);
        Assert.Equal(ExtensionErrorCode.SessionNotFound, result.ErrorCode);
    }

    [Fact]
    public async Task BindAsync_WithUnknownExtension_ShouldReportExtensionNotFound()
    {
        var (bridge, sessions, extensions) = CreateBridge();
        using var _ = bridge;
        using var __ = sessions;
        await using var ___ = extensions;
        var sessionId = sessions.OpenDocument(CreateWordFile("bind.docx"));

        var result = await bridge.BindAsync(sessionId, "no-such-extension", "pdf",
            ConversionOptions.WithoutAHost(), SessionIdentity.GetAnonymous());

        Assert.False(result.IsSuccess);
        Assert.Equal(ExtensionErrorCode.ExtensionNotFound, result.ErrorCode);
    }

    [Fact]
    public async Task BindAsync_AfterDispose_ShouldRefuse()
    {
        var (bridge, sessions, extensions) = CreateBridge();
        using var __ = sessions;
        await using var ___ = extensions;
        var sessionId = sessions.OpenDocument(CreateWordFile("disposed.docx"));
        bridge.Dispose();

        var result = await bridge.BindAsync(sessionId, "any-extension", "pdf",
            ConversionOptions.WithoutAHost(), SessionIdentity.GetAnonymous());

        Assert.False(result.IsSuccess);
        Assert.Equal(ExtensionErrorCode.ExtensionDisabled, result.ErrorCode);
    }

    [Fact]
    public async Task BindAsync_ForAnotherOwnersSession_ShouldNotSeeIt()
    {
        var (bridge, sessions, extensions) = CreateBridge(CreateSessionConfig(SessionIsolationMode.Group));
        using var _ = bridge;
        using var __ = sessions;
        await using var ___ = extensions;

        var owner = new SessionIdentity { GroupId = "group-a", UserId = "user-a" };
        var stranger = new SessionIdentity { GroupId = "group-b", UserId = "user-b" };
        var sessionId = sessions.OpenDocument(CreateWordFile("owned.docx"), owner);

        var result = await bridge.BindAsync(sessionId, "any-extension", "pdf",
            ConversionOptions.WithoutAHost(), stranger);

        Assert.False(result.IsSuccess);
        Assert.Equal(ExtensionErrorCode.SessionNotFound, result.ErrorCode);
    }

    [Fact]
    public async Task Unbind_WithoutABinding_ShouldReportNothingRemoved()
    {
        var (bridge, sessions, extensions) = CreateBridge();
        using var _ = bridge;
        using var __ = sessions;
        await using var ___ = extensions;

        Assert.False(bridge.Unbind("no-such-session", "no-such-extension"));
        Assert.Equal(0, bridge.UnbindAll("no-such-session"));
        Assert.Equal(0, bridge.RemoveBindingsForExtension("no-such-extension"));
    }

    [Fact]
    public async Task GetBindings_WithoutABinding_ShouldBeEmpty()
    {
        var (bridge, sessions, extensions) = CreateBridge();
        using var _ = bridge;
        using var __ = sessions;
        await using var ___ = extensions;
        var sessionId = sessions.OpenDocument(CreateWordFile("empty.docx"));

        Assert.Empty(bridge.GetBindings(sessionId));
        Assert.Empty(bridge.GetBindingsByExtension("no-such-extension"));
    }

    [Fact]
    public async Task Dispose_ShouldBeIdempotent()
    {
        var (bridge, sessions, extensions) = CreateBridge();
        using var __ = sessions;
        await using var ___ = extensions;

        bridge.Dispose();
        var exception = Record.Exception(() => bridge.Dispose());

        Assert.Null(exception);
    }
}
