using Aspose.Words;
using AsposeMcpServer.Core.Session;
using AsposeMcpServer.Results.Session;
using AsposeMcpServer.Tools.Session;
using AsposeMcpServer.Tools.Word;

namespace AsposeMcpServer.Tests.Integration.Workflows;

/// <summary>
///     Integration tests for Word document workflows.
/// </summary>
[Trait("Category", "Integration")]
[Collection("Workflow")]
public class WordWorkflowTests : IntegrationTestBase
{
    private readonly WordFileTool _fileTool;
    private readonly WordHeaderFooterTool _headerFooterTool;
    private readonly DocumentSessionManager _sessionManager;
    private readonly DocumentSessionTool _sessionTool;
    private readonly WordTableTool _tableTool;
    private readonly WordTextTool _textTool;

    /// <summary>
    ///     Initializes a new instance of the <see cref="WordWorkflowTests" /> class.
    /// </summary>
    public WordWorkflowTests()
    {
        var config = new SessionConfig { Enabled = true, TempDirectory = Path.Combine(TestDir, "temp") };
        _sessionManager = new DocumentSessionManager(config);
        var tempFileManager = new TempFileManager(config);
        _sessionTool = new DocumentSessionTool(_sessionManager, tempFileManager, new StdioSessionIdentityAccessor());
        _textTool = new WordTextTool(_sessionManager);
        _tableTool = new WordTableTool(_sessionManager);
        _headerFooterTool = new WordHeaderFooterTool(_sessionManager);
        _fileTool = new WordFileTool(_sessionManager);
    }

    /// <summary>
    ///     Disposes of test resources.
    /// </summary>
    public override void Dispose()
    {
        _sessionManager.Dispose();
        base.Dispose();
    }

    #region Open-Edit-Save Workflow Tests

    /// <summary>
    ///     Verifies the complete open, edit, and save workflow for Word documents.
    /// </summary>
    [SkippableFact]
    public void Word_OpenEditSave_Workflow()
    {
        SkipInEvaluationMode(AsposeLibraryType.Words, "Text operations have limitations in evaluation mode");

        var originalPath = CreateWordDocument("Test Document Content");
        var openResult = _sessionTool.Execute("open", originalPath);
        var openData = GetResultData<OpenSessionResult>(openResult);

        _textTool.Execute("replace", sessionId: openData.SessionId, find: "Test", replace: "Modified");

        var outputPath = CreateTestFilePath("workflow_output.docx");
        _sessionTool.Execute("save", sessionId: openData.SessionId, outputPath: outputPath);

        var savedContent = ReadWordDocumentContent(outputPath);
        Assert.Contains("Modified", savedContent);

        _sessionTool.Execute("close", sessionId: openData.SessionId);
    }

    #endregion

    #region Table Workflow Tests

    /// <summary>
    ///     Verifies the workflow of inserting a table with content.
    /// </summary>
    [Fact]
    public void Word_InsertTableWithContent_Workflow()
    {
        var originalPath = CreateWordDocument("Document with table below:");
        var openResult = _sessionTool.Execute("open", originalPath);
        var openData = GetResultData<OpenSessionResult>(openResult);

        _tableTool.Execute("create", sessionId: openData.SessionId, rows: 3, columns: 3);

        var outputPath = CreateTestFilePath("table_workflow.docx");
        _sessionTool.Execute("save", sessionId: openData.SessionId, outputPath: outputPath);

        var savedDoc = new Document(outputPath);
        var tables = savedDoc.GetChildNodes(NodeType.Table, true);
        Assert.True(tables.Count > 0);

        _sessionTool.Execute("close", sessionId: openData.SessionId);
    }

    #endregion

    #region Find and Replace Workflow Tests

    /// <summary>
    ///     Verifies the find and replace all workflow.
    /// </summary>
    [SkippableFact]
    public void Word_FindReplaceAll_Workflow()
    {
        SkipInEvaluationMode(AsposeLibraryType.Words, "Text operations have limitations in evaluation mode");

        var originalPath = CreateWordDocument("Hello World. Hello again. Hello once more.");
        var openResult = _sessionTool.Execute("open", originalPath);
        var openData = GetResultData<OpenSessionResult>(openResult);

        _textTool.Execute("replace", sessionId: openData.SessionId, find: "Hello", replace: "Hi");

        var outputPath = CreateTestFilePath("replace_all_workflow.docx");
        _sessionTool.Execute("save", sessionId: openData.SessionId, outputPath: outputPath);

        var savedContent = ReadWordDocumentContent(outputPath);
        Assert.DoesNotContain("Hello", savedContent);
        Assert.Contains("Hi", savedContent);

        _sessionTool.Execute("close", sessionId: openData.SessionId);
    }

    #endregion

    #region Header Footer Workflow Tests

    /// <summary>
    ///     Verifies the workflow of adding header and footer.
    /// </summary>
    [Fact]
    public void Word_AddHeaderFooter_Workflow()
    {
        var originalPath = CreateWordDocument("Document content");
        var openResult = _sessionTool.Execute("open", originalPath);
        var openData = GetResultData<OpenSessionResult>(openResult);

        _headerFooterTool.Execute("set_header_text", sessionId: openData.SessionId, headerCenter: "My Header");

        _headerFooterTool.Execute("set_footer_text", sessionId: openData.SessionId, footerCenter: "Page Footer");

        var outputPath = CreateTestFilePath("header_footer_workflow.docx");
        _sessionTool.Execute("save", sessionId: openData.SessionId, outputPath: outputPath);

        var savedDoc = new Document(outputPath);
        var headerFooters = savedDoc.FirstSection.HeadersFooters;
        Assert.True(headerFooters.Count > 0);

        _sessionTool.Execute("close", sessionId: openData.SessionId);
    }

    #endregion

    #region Merge Documents Workflow Tests

    /// <summary>
    ///     Verifies the workflow of merging multiple documents.
    /// </summary>
    [Fact]
    public void Word_MergeDocuments_Workflow()
    {
        var doc1Path = CreateWordDocument("First document content", "merge_doc1.docx");
        var doc2Path = CreateWordDocument("Second document content", "merge_doc2.docx");
        var doc3Path = CreateWordDocument("Third document content", "merge_doc3.docx");

        var outputPath = CreateTestFilePath("merged_output.docx");
        _fileTool.Execute("merge",
            inputPaths: [doc1Path, doc2Path, doc3Path],
            outputPath: outputPath);

        Assert.True(File.Exists(outputPath));

        var mergedDoc = new Document(outputPath);
        var content = mergedDoc.GetText();
        Assert.Contains("First document", content);
        Assert.Contains("Second document", content);
        Assert.Contains("Third document", content);
    }

    #endregion
}
