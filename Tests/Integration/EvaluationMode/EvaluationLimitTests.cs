using Aspose.Words;
using AsposeMcpServer.Core.Session;
using AsposeMcpServer.Results.Session;
using AsposeMcpServer.Tests.Infrastructure;
using AsposeMcpServer.Tools.Session;
using AsposeMcpServer.Tools.Word;

namespace AsposeMcpServer.Tests.Integration.EvaluationMode;

/// <summary>
///     Integration tests for evaluation mode limitations.
///     These tests verify that operations work correctly both with and without license.
/// </summary>
[Trait("Category", "Integration")]
public class EvaluationLimitTests : TestBase
{
    private readonly DocumentSessionManager _sessionManager;
    private readonly DocumentSessionTool _sessionTool;
    private readonly WordTextTool _textTool;

    /// <summary>
    ///     Initializes a new instance of the <see cref="EvaluationLimitTests" /> class.
    /// </summary>
    public EvaluationLimitTests()
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

    #region Session Operations in Evaluation Mode

    /// <summary>
    ///     Verifies that session list works in evaluation mode.
    /// </summary>
    [Fact]
    public void Evaluation_SessionList_WorksInEvaluationMode()
    {
        var path = CreateWordDocument();
        _sessionTool.Execute("open", path);

        var listResult = _sessionTool.Execute("list");
        var listData = GetResultData<ListSessionsResult>(listResult);

        Assert.NotNull(listData.Sessions);
        Assert.NotEmpty(listData.Sessions);
    }

    #endregion

    #region Word Evaluation Mode Tests

    /// <summary>
    ///     Verifies that Word text replacement works with license (skipped in evaluation mode).
    /// </summary>
    [SkippableFact]
    public void Evaluation_WordTextReplace_RequiresLicense()
    {
        SkipInEvaluationMode(AsposeLibraryType.Words, "Text replacement adds watermark in evaluation mode");

        var path = CreateWordDocument("Hello World");
        var openResult = _sessionTool.Execute("open", path);
        var openData = GetResultData<OpenSessionResult>(openResult);

        _textTool.Execute("replace", sessionId: openData.SessionId, find: "Hello", replace: "Hi");

        var outputPath = CreateTestFilePath("eval_word_replace.docx");
        _sessionTool.Execute("save", sessionId: openData.SessionId, outputPath: outputPath);

        var savedContent = ReadWordDocumentContent(outputPath);
        Assert.Contains("Hi", savedContent);
        Assert.DoesNotContain("Hello", savedContent);

        _sessionTool.Execute("close", sessionId: openData.SessionId);
    }

    /// <summary>
    ///     Verifies that basic Word operations work in evaluation mode.
    /// </summary>
    [Fact]
    public void Evaluation_WordBasicOperations_WorksInEvaluationMode()
    {
        var path = CreateWordDocument();
        var openResult = _sessionTool.Execute("open", path);
        var openData = GetResultData<OpenSessionResult>(openResult);

        Assert.NotNull(openData.SessionId);
        Assert.NotEmpty(openData.SessionId);

        _sessionTool.Execute("close", sessionId: openData.SessionId);
    }

    #endregion

    #region Helper Methods

    private string CreateWordDocument(string content = "Test Content")
    {
        var path = CreateTestFilePath($"word_{Guid.NewGuid()}.docx");
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.Writeln(content);
        doc.Save(path);
        return path;
    }

    private static string ReadWordDocumentContent(string path)
    {
        var doc = new Document(path);
        return doc.GetText();
    }

    #endregion
}
