using AsposeMcpServer.Core;
using AsposeMcpServer.Errors;
using ModelContextProtocol.Protocol;

namespace AsposeMcpServer.Tests.Core;

/// <summary>
///     Tests for the tool-call error filter created by <see cref="HostFactory.CreateErrorDetailFilter" />.
///     Designed, in-repo error messages must pass through verbatim; unexpected BCL/Aspose exception
///     text (which can carry absolute file-system paths and internals) must be replaced by fixed
///     sentinels before it reaches the MCP caller.
/// </summary>
public class HostFactoryErrorFilterTests
{
    private static async Task<string> RunFilterWithException(Exception exception)
    {
        var filter = HostFactory.CreateErrorDetailFilter();
        var handler = filter((_, _) => throw exception);

        var result = await handler(null!, CancellationToken.None);

        Assert.True(result.IsError);
        var block = Assert.IsType<TextContentBlock>(Assert.Single(result.Content));
        return block.Text;
    }

    [Fact]
    public async Task Filter_ArgumentException_PassesDesignedMessageThrough()
    {
        var text = await RunFilterWithException(
            new ArgumentException("paragraphIndex 5 is out of range for story 'Body' (it has 3 paragraphs)."));

        Assert.Equal("paragraphIndex 5 is out of range for story 'Body' (it has 3 paragraphs).", text);
    }

    [Fact]
    public async Task Filter_KeyNotFoundException_PassesDesignedMessageThrough()
    {
        var text = await RunFilterWithException(new KeyNotFoundException("Session not found: sess_123"));

        Assert.Equal("Session not found: sess_123", text);
    }

    [Fact]
    public async Task Filter_SessionDisposedException_PassesDesignedMessageThrough()
    {
        var text = await RunFilterWithException(
            new ObjectDisposedException("DocumentSession", "Session sess_123 has been disposed"));

        Assert.Contains("Session sess_123 has been disposed", text);
    }

    [Fact]
    public async Task Filter_UnexpectedIoException_DoesNotLeakPaths()
    {
        var text = await RunFilterWithException(
            new IOException(@"The process cannot access the file 'C:\server\internal\temp\abc123.tmp'."));

        Assert.DoesNotContain(@"C:\server", text);
        Assert.Equal(ErrorMessageBuilder.ProcessingFailed(), text);
    }

    [Fact]
    public async Task Filter_BclFileNotFoundException_DoesNotLeakPath()
    {
        var text = await RunFilterWithException(
            new FileNotFoundException(@"Could not find file 'C:\server\secret\doc.docx'.",
                @"C:\server\secret\doc.docx"));

        Assert.DoesNotContain(@"C:\server", text);
        Assert.Equal("The specified file was not found.", text);
    }

    [Fact]
    public async Task Filter_RawUnauthorizedAccessException_DoesNotLeakPath()
    {
        var text = await RunFilterWithException(
            new UnauthorizedAccessException(@"Access to the path 'C:\server\secret\doc.docx' is denied."));

        Assert.DoesNotContain(@"C:\server", text);
    }

    [Fact]
    public async Task Filter_PasswordSentinelUnauthorizedAccessException_PassesThrough()
    {
        var text = await RunFilterWithException(
            new UnauthorizedAccessException(ErrorMessageBuilder.InvalidPassword()));

        Assert.Equal(ErrorMessageBuilder.InvalidPassword(), text);
    }
}
