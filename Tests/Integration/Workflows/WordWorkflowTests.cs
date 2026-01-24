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
        var config = new SessionConfig { Enabled = true };
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

        // Step 1: Create and open document
        var originalPath = CreateWordDocument("Test Document Content");
        var openResult = _sessionTool.Execute("open", originalPath);
        var openData = GetResultData<OpenSessionResult>(openResult);

        // Step 2: Edit document - replace text
        _textTool.Execute("replace", sessionId: openData.SessionId, find: "Test", replace: "Modified");

        // Step 3: Save document
        var outputPath = CreateTestFilePath("workflow_output.docx");
        _sessionTool.Execute("save", sessionId: openData.SessionId, outputPath: outputPath);

        // Step 4: Verify changes persisted
        var savedContent = ReadWordDocumentContent(outputPath);
        Assert.Contains("Modified", savedContent);

        // Step 5: Close session
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
        // Step 1: Create and open document
        var originalPath = CreateWordDocument("Document with table below:");
        var openResult = _sessionTool.Execute("open", originalPath);
        var openData = GetResultData<OpenSessionResult>(openResult);

        // Step 2: Create table
        _tableTool.Execute("create", sessionId: openData.SessionId, rows: 3, columns: 3);

        // Step 3: Save and verify
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

        // Step 1: Create document with multiple occurrences
        var originalPath = CreateWordDocument("Hello World. Hello again. Hello once more.");
        var openResult = _sessionTool.Execute("open", originalPath);
        var openData = GetResultData<OpenSessionResult>(openResult);

        // Step 2: Replace all occurrences
        _textTool.Execute("replace", sessionId: openData.SessionId, find: "Hello", replace: "Hi");

        // Step 3: Save and verify all occurrences replaced
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
        // Step 1: Create and open document
        var originalPath = CreateWordDocument("Document content");
        var openResult = _sessionTool.Execute("open", originalPath);
        var openData = GetResultData<OpenSessionResult>(openResult);

        // Step 2: Add header text
        _headerFooterTool.Execute("set_header_text", sessionId: openData.SessionId, headerCenter: "My Header");

        // Step 3: Add footer text
        _headerFooterTool.Execute("set_footer_text", sessionId: openData.SessionId, footerCenter: "Page Footer");

        // Step 4: Save and verify
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
        // Step 1: Create multiple documents
        var doc1Path = CreateWordDocument("First document content", "merge_doc1.docx");
        var doc2Path = CreateWordDocument("Second document content", "merge_doc2.docx");
        var doc3Path = CreateWordDocument("Third document content", "merge_doc3.docx");

        // Step 2: Merge documents
        var outputPath = CreateTestFilePath("merged_output.docx");
        _fileTool.Execute("merge",
            inputPaths: [doc1Path, doc2Path, doc3Path],
            outputPath: outputPath);

        // Step 3: Verify merged document
        Assert.True(File.Exists(outputPath));

        var mergedDoc = new Document(outputPath);
        var content = mergedDoc.GetText();
        Assert.Contains("First document", content);
        Assert.Contains("Second document", content);
        Assert.Contains("Third document", content);
    }

    #endregion
}
