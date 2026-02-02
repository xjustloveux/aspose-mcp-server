using Aspose.Words;
using AsposeMcpServer.Core.Session;
using AsposeMcpServer.Results.Session;
using AsposeMcpServer.Tests.Infrastructure;
using AsposeMcpServer.Tools.Session;

namespace AsposeMcpServer.Tests.Integration.ErrorHandling;

/// <summary>
///     Integration tests for system-level error handling.
/// </summary>
[Trait("Category", "Integration")]
public class SystemErrorTests : TestBase
{
    private readonly DocumentSessionManager _sessionManager;
    private readonly DocumentSessionTool _sessionTool;

    /// <summary>
    ///     Initializes a new instance of the <see cref="SystemErrorTests" /> class.
    /// </summary>
    public SystemErrorTests()
    {
        var config = new SessionConfig { Enabled = true, TempDirectory = Path.Combine(TestDir, "temp") };
        _sessionManager = new DocumentSessionManager(config);
        var tempFileManager = new TempFileManager(config);
        _sessionTool = new DocumentSessionTool(_sessionManager, tempFileManager, new StdioSessionIdentityAccessor());
    }

    /// <summary>
    ///     Disposes of test resources.
    /// </summary>
    public override void Dispose()
    {
        _sessionManager.Dispose();
        base.Dispose();
    }

    #region Timeout Simulation Tests

    /// <summary>
    ///     Verifies that very large documents can still be processed.
    /// </summary>
    [Fact]
    public void System_LargeDocument_ProcessesSuccessfully()
    {
        var path = CreateTestFilePath("large_doc.docx");
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        for (var i = 0; i < 100; i++)
            builder.Writeln($"Paragraph {i}: This is some test content to make the document larger.");

        doc.Save(path);

        var openResult = _sessionTool.Execute("open", path);
        var openData = GetResultData<OpenSessionResult>(openResult);

        Assert.NotNull(openData.SessionId);

        var outputPath = CreateTestFilePath("large_doc_output.docx");
        _sessionTool.Execute("save", sessionId: openData.SessionId, outputPath: outputPath);

        Assert.True(File.Exists(outputPath));

        _sessionTool.Execute("close", sessionId: openData.SessionId);
    }

    #endregion

    #region Helper Methods

    private string CreateWordDocument(string content = "Test Content", string? fileName = null)
    {
        var path = CreateTestFilePath(fileName ?? $"word_{Guid.NewGuid()}.docx");
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.Writeln(content);
        doc.Save(path);
        return path;
    }

    #endregion

    #region Disk Space Tests

    /// <summary>
    ///     Verifies that saving to an invalid path throws exception.
    /// </summary>
    [Fact]
    public void System_InvalidOutputPath_ThrowsException()
    {
        var path = CreateWordDocument();
        var openResult = _sessionTool.Execute("open", path);
        var openData = GetResultData<OpenSessionResult>(openResult);

        // Try to save to an invalid path (non-existent drive on Windows)
        var invalidPath = Path.Combine("Z:", "nonexistent", "folder", "output.docx");

        Assert.ThrowsAny<Exception>(() =>
            _sessionTool.Execute("save", sessionId: openData.SessionId, outputPath: invalidPath));
    }

    /// <summary>
    ///     Verifies that saving to a read-only file throws exception.
    /// </summary>
    [Fact]
    public void System_ReadOnlyFile_ThrowsException()
    {
        var path = CreateWordDocument();
        var openResult = _sessionTool.Execute("open", path);
        var openData = GetResultData<OpenSessionResult>(openResult);

        var readOnlyFile = CreateTestFilePath("readonly_output.docx");
        File.WriteAllText(readOnlyFile, "existing content");
        File.SetAttributes(readOnlyFile, FileAttributes.ReadOnly);

        try
        {
            Assert.ThrowsAny<Exception>(() =>
                _sessionTool.Execute("save", sessionId: openData.SessionId, outputPath: readOnlyFile));
        }
        finally
        {
            File.SetAttributes(readOnlyFile, FileAttributes.Normal);
        }
    }

    #endregion

    #region Resource Limit Tests

    /// <summary>
    ///     Verifies that opening many sessions works until limit is reached.
    /// </summary>
    [Fact]
    public void System_ManySessions_HandlesCorrectly()
    {
        var sessions = new List<string>();

        try
        {
            for (var i = 0; i < 10; i++)
            {
                var path = CreateWordDocument(fileName: $"multi_session_{i}.docx");
                var openResult = _sessionTool.Execute("open", path);
                var openData = GetResultData<OpenSessionResult>(openResult);
                sessions.Add(openData.SessionId);
            }

            Assert.Equal(10, sessions.Count);

            var listResult = _sessionTool.Execute("list");
            var listData = GetResultData<ListSessionsResult>(listResult);
            Assert.True(listData.Sessions.Count >= 10);
        }
        finally
        {
            foreach (var sessionId in sessions)
                try
                {
                    _sessionTool.Execute("close", sessionId: sessionId);
                }
                catch
                {
                    // Ignore cleanup errors
                }
        }
    }

    /// <summary>
    ///     Verifies that operations on closed sessions throw KeyNotFoundException.
    /// </summary>
    [Fact]
    public void System_OperationOnClosedSession_ThrowsKeyNotFoundException()
    {
        var path = CreateWordDocument();
        var openResult = _sessionTool.Execute("open", path);
        var openData = GetResultData<OpenSessionResult>(openResult);
        var sessionId = openData.SessionId;

        _sessionTool.Execute("close", sessionId: sessionId);

        Assert.Throws<KeyNotFoundException>(() =>
            _sessionTool.Execute("save", sessionId: sessionId, outputPath: CreateTestFilePath("output.docx")));
    }

    #endregion
}
