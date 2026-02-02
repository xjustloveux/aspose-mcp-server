using AsposeMcpServer.Core.Session;
using AsposeMcpServer.Results.Session;
using AsposeMcpServer.Tools.Session;

namespace AsposeMcpServer.Tests.Integration.Session;

/// <summary>
///     Integration tests for document session lifecycle management.
/// </summary>
[Trait("Category", "Integration")]
[Collection("Session Integration")]
public class SessionLifecycleTests : IntegrationTestBase
{
    private readonly DocumentSessionManager _sessionManager;
    private readonly DocumentSessionTool _sessionTool;

    /// <summary>
    ///     Initializes a new instance of the <see cref="SessionLifecycleTests" /> class.
    /// </summary>
    public SessionLifecycleTests()
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

    #region Open Session Tests

    /// <summary>
    ///     Verifies that opening a Word document creates a session.
    /// </summary>
    [Fact]
    public void Session_OpenWord_CreatesSession()
    {
        var path = CreateWordDocument();

        var result = _sessionTool.Execute("open", path);
        var data = GetResultData<OpenSessionResult>(result);

        Assert.NotNull(data.SessionId);
        Assert.NotEmpty(data.SessionId);
    }

    /// <summary>
    ///     Verifies that opening an Excel document creates a session.
    /// </summary>
    [Fact]
    public void Session_OpenExcel_CreatesSession()
    {
        var path = CreateExcelDocument();

        var result = _sessionTool.Execute("open", path);
        var data = GetResultData<OpenSessionResult>(result);

        Assert.NotNull(data.SessionId);
        Assert.NotEmpty(data.SessionId);
    }

    /// <summary>
    ///     Verifies that opening a PowerPoint document creates a session.
    /// </summary>
    [Fact]
    public void Session_OpenPowerPoint_CreatesSession()
    {
        var path = CreatePowerPointDocument();

        var result = _sessionTool.Execute("open", path);
        var data = GetResultData<OpenSessionResult>(result);

        Assert.NotNull(data.SessionId);
        Assert.NotEmpty(data.SessionId);
    }

    /// <summary>
    ///     Verifies that opening a PDF document creates a session.
    /// </summary>
    [Fact]
    public void Session_OpenPdf_CreatesSession()
    {
        var path = CreatePdfDocument();

        var result = _sessionTool.Execute("open", path);
        var data = GetResultData<OpenSessionResult>(result);

        Assert.NotNull(data.SessionId);
        Assert.NotEmpty(data.SessionId);
    }

    /// <summary>
    ///     Verifies that opening a document returns a unique session ID.
    /// </summary>
    [Fact]
    public void Session_Open_ReturnsUniqueSessionId()
    {
        var path1 = CreateWordDocument(fileName: "doc1.docx");
        var path2 = CreateWordDocument(fileName: "doc2.docx");

        var result1 = _sessionTool.Execute("open", path1);
        var result2 = _sessionTool.Execute("open", path2);

        var data1 = GetResultData<OpenSessionResult>(result1);
        var data2 = GetResultData<OpenSessionResult>(result2);

        Assert.NotEqual(data1.SessionId, data2.SessionId);
    }

    #endregion

    #region Close Session Tests

    /// <summary>
    ///     Verifies that closing a session releases resources.
    /// </summary>
    [Fact]
    public void Session_Close_ReleasesResources()
    {
        var path = CreateWordDocument();
        var openResult = _sessionTool.Execute("open", path);
        var openData = GetResultData<OpenSessionResult>(openResult);
        var sessionId = openData.SessionId;

        var closeResult = _sessionTool.Execute("close", sessionId: sessionId);

        Assert.NotNull(closeResult);
        Assert.Throws<KeyNotFoundException>(() => _sessionManager.GetDocument<object>(sessionId));
    }

    /// <summary>
    ///     Verifies that closing an invalid session throws KeyNotFoundException.
    /// </summary>
    [Fact]
    public void Session_CloseInvalidId_ThrowsKeyNotFoundException()
    {
        Assert.Throws<KeyNotFoundException>(() => _sessionTool.Execute("close", sessionId: "invalid-session-id"));
    }

    #endregion

    #region List Sessions Tests

    /// <summary>
    ///     Verifies that listing sessions returns all open sessions.
    /// </summary>
    [Fact]
    public void Session_List_ReturnsOpenSessions()
    {
        var path1 = CreateWordDocument(fileName: "list1.docx");
        var path2 = CreateExcelDocument("list2.xlsx");

        _sessionTool.Execute("open", path1);
        _sessionTool.Execute("open", path2);

        var listResult = _sessionTool.Execute("list");
        var listData = GetResultData<ListSessionsResult>(listResult);

        Assert.True(listData.Sessions.Count >= 2);
    }

    /// <summary>
    ///     Verifies that listing sessions returns correct document types.
    /// </summary>
    [Fact]
    public void Session_List_ReturnsCorrectDocumentTypes()
    {
        var wordPath = CreateWordDocument(fileName: "type_word.docx");
        var excelPath = CreateExcelDocument("type_excel.xlsx");

        _sessionTool.Execute("open", wordPath);
        _sessionTool.Execute("open", excelPath);

        var listResult = _sessionTool.Execute("list");
        var listData = GetResultData<ListSessionsResult>(listResult);

        Assert.Contains(listData.Sessions, s => s.DocumentType == "word");
        Assert.Contains(listData.Sessions, s => s.DocumentType == "excel");
    }

    #endregion
}
