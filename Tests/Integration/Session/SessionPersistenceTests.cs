using Aspose.Cells;
using AsposeMcpServer.Core.Session;
using AsposeMcpServer.Results.Session;
using AsposeMcpServer.Tools.Session;
using AsposeMcpServer.Tools.Word;

namespace AsposeMcpServer.Tests.Integration.Session;

/// <summary>
///     Integration tests for document session persistence.
/// </summary>
[Trait("Category", "Integration")]
[Collection("Session Integration")]
public class SessionPersistenceTests : IntegrationTestBase
{
    private readonly DocumentSessionManager _sessionManager;
    private readonly DocumentSessionTool _sessionTool;
    private readonly WordTextTool _textTool;

    /// <summary>
    ///     Initializes a new instance of the <see cref="SessionPersistenceTests" /> class.
    /// </summary>
    public SessionPersistenceTests()
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

    #region Save Session Tests

    /// <summary>
    ///     Verifies that saving a session persists changes to the file.
    /// </summary>
    [SkippableFact]
    public void Session_Save_PersistsChanges()
    {
        SkipInEvaluationMode(AsposeLibraryType.Words, "Text replacement has limitations in evaluation mode");

        var originalPath = CreateWordDocument();
        var outputPath = CreateTestFilePath("saved_output.docx");

        var openResult = _sessionTool.Execute("open", originalPath);
        var openData = GetResultData<OpenSessionResult>(openResult);

        _textTool.Execute("replace", sessionId: openData.SessionId, find: "Test", replace: "Modified");

        _sessionTool.Execute("save", sessionId: openData.SessionId, outputPath: outputPath);

        Assert.True(File.Exists(outputPath));

        var savedContent = ReadWordDocumentContent(outputPath);
        Assert.Contains("Modified", savedContent);
    }

    /// <summary>
    ///     Verifies that save as creates a new file without affecting the original.
    /// </summary>
    [SkippableFact]
    public void Session_SaveAs_CreatesNewFile()
    {
        SkipInEvaluationMode(AsposeLibraryType.Words, "Text replacement has limitations in evaluation mode");

        var originalPath = CreateWordDocument();
        var newPath = CreateTestFilePath("save_as_output.docx");

        var openResult = _sessionTool.Execute("open", originalPath);
        var openData = GetResultData<OpenSessionResult>(openResult);

        _textTool.Execute("replace", sessionId: openData.SessionId, find: "Test", replace: "SavedAs");

        _sessionTool.Execute("save", sessionId: openData.SessionId, outputPath: newPath);

        Assert.True(File.Exists(newPath));

        var originalContent = ReadWordDocumentContent(originalPath);
        Assert.Contains("Test", originalContent);
        Assert.DoesNotContain("SavedAs", originalContent);

        var newContent = ReadWordDocumentContent(newPath);
        Assert.Contains("SavedAs", newContent);
    }

    /// <summary>
    ///     Verifies that Excel sessions can be saved.
    /// </summary>
    [Fact]
    public void Session_SaveExcel_PersistsChanges()
    {
        var originalPath = CreateExcelDocument();
        var outputPath = CreateTestFilePath("saved_excel.xlsx");

        var openResult = _sessionTool.Execute("open", originalPath);
        var openData = GetResultData<OpenSessionResult>(openResult);

        var workbook = _sessionManager.GetDocument<Workbook>(openData.SessionId);
        workbook.Worksheets[0].Cells["B1"].Value = "Modified";

        _sessionTool.Execute("save", sessionId: openData.SessionId, outputPath: outputPath);

        Assert.True(File.Exists(outputPath));

        var savedValue = ReadExcelCellValue(outputPath, "B1");
        Assert.Equal("Modified", savedValue);
    }

    #endregion
}
