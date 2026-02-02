using AsposeMcpServer.Core.Session;
using AsposeMcpServer.Tests.Infrastructure;
using AsposeMcpServer.Tools.Session;

namespace AsposeMcpServer.Tests.Integration.ErrorHandling;

/// <summary>
///     Integration tests for document-related error handling.
/// </summary>
[Trait("Category", "Integration")]
public class DocumentErrorTests : TestBase
{
    private readonly DocumentSessionManager _sessionManager;
    private readonly DocumentSessionTool _sessionTool;

    /// <summary>
    ///     Initializes a new instance of the <see cref="DocumentErrorTests" /> class.
    /// </summary>
    public DocumentErrorTests()
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

    #region Unsupported Format Tests

    /// <summary>
    ///     Verifies that opening an unsupported file format throws exception.
    /// </summary>
    [Fact]
    public void Document_UnsupportedFormat_ThrowsException()
    {
        var unsupportedPath = CreateTestFilePath("unsupported.xyz");
        File.WriteAllText(unsupportedPath, "Content");

        Assert.ThrowsAny<Exception>(() => _sessionTool.Execute("open", unsupportedPath));
    }

    #endregion

    #region File Not Found Tests

    /// <summary>
    ///     Verifies that opening a non-existent file throws FileNotFoundException.
    /// </summary>
    [Fact]
    public void Document_FileNotFound_ThrowsException()
    {
        var nonExistentPath = Path.Combine(TestDir, "non_existent_file.docx");

        Assert.ThrowsAny<Exception>(() => _sessionTool.Execute("open", nonExistentPath));
    }

    /// <summary>
    ///     Verifies that opening a non-existent Excel file throws exception.
    /// </summary>
    [Fact]
    public void Document_ExcelFileNotFound_ThrowsException()
    {
        var nonExistentPath = Path.Combine(TestDir, "non_existent_file.xlsx");

        Assert.ThrowsAny<Exception>(() => _sessionTool.Execute("open", nonExistentPath));
    }

    /// <summary>
    ///     Verifies that opening a non-existent PowerPoint file throws exception.
    /// </summary>
    [Fact]
    public void Document_PowerPointFileNotFound_ThrowsException()
    {
        var nonExistentPath = Path.Combine(TestDir, "non_existent_file.pptx");

        Assert.ThrowsAny<Exception>(() => _sessionTool.Execute("open", nonExistentPath));
    }

    /// <summary>
    ///     Verifies that opening a non-existent PDF file throws exception.
    /// </summary>
    [Fact]
    public void Document_PdfFileNotFound_ThrowsException()
    {
        var nonExistentPath = Path.Combine(TestDir, "non_existent_file.pdf");

        Assert.ThrowsAny<Exception>(() => _sessionTool.Execute("open", nonExistentPath));
    }

    #endregion

    #region Corrupted File Tests

    /// <summary>
    ///     Verifies that opening a corrupted Excel file throws exception.
    /// </summary>
    [Fact]
    public void Document_CorruptedExcel_ThrowsException()
    {
        var corruptedPath = CreateTestFilePath("corrupted.xlsx");
        File.WriteAllText(corruptedPath, "This is not a valid Excel file");

        Assert.ThrowsAny<Exception>(() => _sessionTool.Execute("open", corruptedPath));
    }

    /// <summary>
    ///     Verifies that opening a corrupted PowerPoint file throws exception.
    /// </summary>
    [Fact]
    public void Document_CorruptedPowerPoint_ThrowsException()
    {
        var corruptedPath = CreateTestFilePath("corrupted.pptx");
        File.WriteAllText(corruptedPath, "This is not a valid PowerPoint file");

        Assert.ThrowsAny<Exception>(() => _sessionTool.Execute("open", corruptedPath));
    }

    /// <summary>
    ///     Verifies that opening a corrupted PDF file throws exception.
    /// </summary>
    [Fact]
    public void Document_CorruptedPdf_ThrowsException()
    {
        var corruptedPath = CreateTestFilePath("corrupted.pdf");
        File.WriteAllText(corruptedPath, "This is not a valid PDF file");

        Assert.ThrowsAny<Exception>(() => _sessionTool.Execute("open", corruptedPath));
    }

    #endregion

    #region Invalid Session Tests

    /// <summary>
    ///     Verifies that using an invalid session ID throws KeyNotFoundException.
    /// </summary>
    [Fact]
    public void Document_InvalidSessionId_ThrowsKeyNotFoundException()
    {
        Assert.Throws<KeyNotFoundException>(() =>
            _sessionTool.Execute("save", sessionId: "invalid-session-12345"));
    }

    /// <summary>
    ///     Verifies that closing an invalid session ID throws KeyNotFoundException.
    /// </summary>
    [Fact]
    public void Document_CloseInvalidSession_ThrowsKeyNotFoundException()
    {
        Assert.Throws<KeyNotFoundException>(() =>
            _sessionTool.Execute("close", sessionId: "invalid-session-67890"));
    }

    #endregion
}
