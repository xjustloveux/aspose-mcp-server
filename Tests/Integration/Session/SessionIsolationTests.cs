using Aspose.Cells;
using Aspose.Words;
using AsposeMcpServer.Core.Session;
using AsposeMcpServer.Results.Session;
using AsposeMcpServer.Tools.Session;
using AsposeMcpServer.Tools.Word;

namespace AsposeMcpServer.Tests.Integration.Session;

/// <summary>
///     Integration tests for document session isolation.
/// </summary>
[Trait("Category", "Integration")]
[Collection("Session Integration")]
public class SessionIsolationTests : IntegrationTestBase
{
    private readonly DocumentSessionManager _sessionManager;
    private readonly DocumentSessionTool _sessionTool;
    private readonly WordTextTool _textTool;

    /// <summary>
    ///     Initializes a new instance of the <see cref="SessionIsolationTests" /> class.
    /// </summary>
    public SessionIsolationTests()
    {
        var config = new SessionConfig { Enabled = true, TempDirectory = Path.Combine(TestDir, "temp") };
        _sessionManager = new DocumentSessionManager(config);
        var tempFileManager = new TempFileManager(config);
        _sessionTool = new DocumentSessionTool(_sessionManager, tempFileManager, new StdioSessionIdentityAccessor());
        _textTool = new WordTextTool(_sessionManager);
    }

    /// <summary>
    ///     Disposes of test resources.
    /// </summary>
    public override void Dispose()
    {
        _sessionManager.Dispose();
        base.Dispose();
    }

    #region Session Isolation Tests

    /// <summary>
    ///     Verifies that the same document can have multiple independent sessions.
    /// </summary>
    [Fact]
    public void Session_SameDocument_DifferentSessions()
    {
        var path = CreateWordDocument();

        var result1 = _sessionTool.Execute("open", path);
        var result2 = _sessionTool.Execute("open", path);

        var data1 = GetResultData<OpenSessionResult>(result1);
        var data2 = GetResultData<OpenSessionResult>(result2);

        Assert.NotEqual(data1.SessionId, data2.SessionId);
    }

    /// <summary>
    ///     Verifies that modifications in one session do not affect another.
    /// </summary>
    [SkippableFact]
    public void Session_ModifyInOne_NotAffectOther()
    {
        SkipInEvaluationMode(AsposeLibraryType.Words, "Text operations have limitations in evaluation mode");

        var path = CreateWordDocument();

        var result1 = _sessionTool.Execute("open", path);
        var result2 = _sessionTool.Execute("open", path);

        var data1 = GetResultData<OpenSessionResult>(result1);
        var data2 = GetResultData<OpenSessionResult>(result2);

        _textTool.Execute("replace", sessionId: data1.SessionId, find: "Test", replace: "Modified");

        var doc2 = _sessionManager.GetDocument<Document>(data2.SessionId);
        var text2 = doc2.GetText();

        Assert.Contains("Test", text2);
        Assert.DoesNotContain("Modified", text2);
    }

    /// <summary>
    ///     Verifies that different document types maintain separate sessions.
    /// </summary>
    [Fact]
    public void Session_DifferentDocumentTypes_Isolated()
    {
        var wordPath = CreateWordDocument();
        var excelPath = CreateExcelDocument();

        var wordResult = _sessionTool.Execute("open", wordPath);
        var excelResult = _sessionTool.Execute("open", excelPath);

        var wordData = GetResultData<OpenSessionResult>(wordResult);
        var excelData = GetResultData<OpenSessionResult>(excelResult);

        Assert.NotEqual(wordData.SessionId, excelData.SessionId);

        var wordDoc = _sessionManager.GetDocument<Document>(wordData.SessionId);
        var excelDoc = _sessionManager.GetDocument<Workbook>(excelData.SessionId);

        Assert.NotNull(wordDoc);
        Assert.NotNull(excelDoc);
    }

    #endregion
}
