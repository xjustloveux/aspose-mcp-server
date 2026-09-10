using AsposeMcpServer.Core.Extension;
using AsposeMcpServer.Core.Session;
using AsposeMcpServer.Results;
using AsposeMcpServer.Results.Extension;
using AsposeMcpServer.Tests.Infrastructure;
using AsposeMcpServer.Tools.Extension;
using Microsoft.Extensions.Logging.Abstractions;

namespace AsposeMcpServer.Tests.Tools.Extension;

/// <summary>
///     Covers TEST-07 for <see cref="ExtensionTool" />, which sat at a few percent of its lines
///     covered. This is the caller-facing surface for extension bindings, so its dispatch table and
///     its refusals decide what an unreachable extension looks like to a client: an unknown
///     operation must be rejected, and every operation naming a missing session or extension must
///     report that rather than fall through.
/// </summary>
public class ExtensionToolTests : TestBase
{
    /// <summary>Builds a tool over live managers with no extensions registered.</summary>
    /// <returns>The tool plus everything it owns, for the caller to dispose.</returns>
    private (ExtensionTool Tool, ExtensionSessionBridge Bridge, DocumentSessionManager Sessions,
        ExtensionManager Extensions) CreateTool()
    {
        var extensionConfig = new ExtensionConfig { Enabled = true, TempDirectory = TestDir };
        var sessionConfig = new SessionConfig
        {
            Enabled = true,
            IdleTimeoutMinutes = 0,
            TempDirectory = TestDir,
            IsolationMode = SessionIsolationMode.None
        };

        var sessions = new DocumentSessionManager(sessionConfig);
        var snapshots = new SnapshotManager(extensionConfig, NullLogger<SnapshotManager>.Instance);
        var extensions = new ExtensionManager(extensionConfig, snapshots,
            NullLoggerFactory.Instance, NullLogger<ExtensionManager>.Instance);
        var bridge = new ExtensionSessionBridge(extensionConfig, sessions, extensions,
            NullLogger<ExtensionSessionBridge>.Instance);
        var tool = new ExtensionTool(extensions, bridge, new FixedIdentityAccessor());
        return (tool, bridge, sessions, extensions);
    }

    [Fact]
    public async Task Execute_WithUnknownOperation_ShouldBeRejected()
    {
        var (tool, bridge, sessions, extensions) = CreateTool();
        using var _ = bridge;
        using var __ = sessions;
        await using var ___ = extensions;

        await Assert.ThrowsAnyAsync<ArgumentException>(() => tool.ExecuteAsync("not-an-operation"));
    }

    [Fact]
    public async Task List_WithNoExtensionsRegistered_ShouldReturnAnEmptyList()
    {
        var (tool, bridge, sessions, extensions) = CreateTool();
        using var _ = bridge;
        using var __ = sessions;
        await using var ___ = extensions;

        var raw = await tool.ExecuteAsync("list");

        var data = ((FinalizedResult<ListExtensionsResult>)raw).Data;
        Assert.Empty(data.Extensions);
    }

    [Fact]
    public async Task Bind_WithUnknownSession_ShouldReportFailureWithoutThrowing()
    {
        var (tool, bridge, sessions, extensions) = CreateTool();
        using var _ = bridge;
        using var __ = sessions;
        await using var ___ = extensions;

        var raw = await tool.ExecuteAsync("bind",
            "no-such-session", "no-such-extension", "pdf");

        var data = ((FinalizedResult<BindExtensionResult>)raw).Data;
        Assert.False(data.Success);
        Assert.Equal(ExtensionErrorCode.SessionNotFound, data.ErrorCode);
    }

    [Fact]
    public async Task Bindings_ForASessionWithNone_ShouldReturnAnEmptyList()
    {
        var (tool, bridge, sessions, extensions) = CreateTool();
        using var _ = bridge;
        using var __ = sessions;
        await using var ___ = extensions;

        var raw = await tool.ExecuteAsync("bindings", "no-such-session");

        var data = ((FinalizedResult<ExtensionBindingsResult>)raw).Data;
        Assert.Empty(data.Bindings);
    }

    [Fact]
    public async Task Unbind_WithNoBinding_ShouldReportThatNothingWasRemoved()
    {
        var (tool, bridge, sessions, extensions) = CreateTool();
        using var _ = bridge;
        using var __ = sessions;
        await using var ___ = extensions;

        var raw = await tool.ExecuteAsync("unbind", "no-such-session");

        var data = ((FinalizedResult<UnbindExtensionResult>)raw).Data;
        Assert.Equal(0, data.UnboundCount);
    }

    [Fact]
    public async Task Command_ForAnUnknownExtension_ShouldReportFailure()
    {
        var (tool, bridge, sessions, extensions) = CreateTool();
        using var _ = bridge;
        using var __ = sessions;
        await using var ___ = extensions;

        var raw = await tool.ExecuteAsync("command",
            "no-such-session", "no-such-extension", commandType: "ping");

        var data = ((FinalizedResult<SendCommandResult>)raw).Data;
        Assert.False(data.Success);
        Assert.Equal("no-such-extension", data.ExtensionId);
    }

    [Fact]
    public async Task OperationName_ShouldBeCaseInsensitive()
    {
        var (tool, bridge, sessions, extensions) = CreateTool();
        using var _ = bridge;
        using var __ = sessions;
        await using var ___ = extensions;

        var raw = await tool.ExecuteAsync("LIST");

        var data = ((FinalizedResult<ListExtensionsResult>)raw).Data;
        Assert.Empty(data.Extensions);
    }

    /// <summary>Identity accessor returning a fixed identity, standing in for the HTTP one.</summary>
    private sealed class FixedIdentityAccessor : ISessionIdentityAccessor
    {
        public SessionIdentity GetCurrentIdentity()
        {
            return SessionIdentity.GetAnonymous();
        }
    }
}
