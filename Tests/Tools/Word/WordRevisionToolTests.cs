using Aspose.Words;
using AsposeMcpServer.Tests.Helpers;
using AsposeMcpServer.Tools.Word;

namespace AsposeMcpServer.Tests.Tools.Word;

public class WordRevisionToolTests : WordTestBase
{
    private readonly WordRevisionTool _tool;

    public WordRevisionToolTests()
    {
        _tool = new WordRevisionTool(SessionManager);
    }

    #region General Tests

    [Fact]
    public void GetRevisions_ShouldReturnRevisions()
    {
        var docPath = CreateWordDocument("test_get_revisions.docx");
        var doc = new Document(docPath);
        doc.StartTrackRevisions("Test Author");
        var builder = new DocumentBuilder(doc);
        builder.Writeln("Original text");
        builder.Writeln("Modified text");
        doc.StopTrackRevisions();
        doc.Save(docPath);
        var result = _tool.Execute("get_revisions", docPath);
        Assert.NotNull(result);
        Assert.NotEmpty(result);
        Assert.Contains("\"revisions\"", result); // JSON format
        Assert.Contains("\"index\"", result); // Has index property
    }

    [Fact]
    public void GetRevisions_WithNoRevisions_ShouldReturnZeroCount()
    {
        var docPath = CreateWordDocumentWithContent("test_no_revisions.docx", "Plain text");
        var result = _tool.Execute("get_revisions", docPath);
        Assert.Contains("\"count\": 0", result); // JSON format
    }

    [Fact]
    public void AcceptAllRevisions_ShouldAcceptAll()
    {
        var docPath = CreateWordDocument("test_accept_all_revisions.docx");
        var doc = new Document(docPath);
        doc.StartTrackRevisions("Test Author");
        var builder = new DocumentBuilder(doc);
        builder.Writeln("Text with revision");
        doc.StopTrackRevisions();
        doc.Save(docPath);

        var revisionsBefore = doc.Revisions.Count;
        Assert.True(revisionsBefore > 0, "Document should have revisions before accepting");

        var outputPath = CreateTestFilePath("test_accept_all_revisions_output.docx");
        var result = _tool.Execute("accept_all", docPath, outputPath: outputPath);
        var resultDoc = new Document(outputPath);
        Assert.Equal(0, resultDoc.Revisions.Count);
        Assert.Contains("Accepted", result);
    }

    [Fact]
    public void RejectAllRevisions_ShouldRejectAll()
    {
        var docPath = CreateWordDocument("test_reject_all_revisions.docx");
        var doc = new Document(docPath);
        doc.StartTrackRevisions("Test Author");
        var builder = new DocumentBuilder(doc);
        builder.Writeln("Text with revision");
        doc.StopTrackRevisions();
        doc.Save(docPath);

        var outputPath = CreateTestFilePath("test_reject_all_revisions_output.docx");
        var result = _tool.Execute("reject_all", docPath, outputPath: outputPath);
        var resultDoc = new Document(outputPath);
        Assert.Equal(0, resultDoc.Revisions.Count);
        Assert.Contains("Rejected", result);
    }

    [Fact]
    public void ManageRevision_Accept_ShouldAcceptSpecificRevision()
    {
        var docPath = CreateWordDocument("test_manage_accept.docx");
        var doc = new Document(docPath);
        doc.StartTrackRevisions("Test Author");
        var builder = new DocumentBuilder(doc);
        builder.Writeln("First revision");
        builder.Writeln("Second revision");
        doc.StopTrackRevisions();
        doc.Save(docPath);

        var outputPath = CreateTestFilePath("test_manage_accept_output.docx");
        var result = _tool.Execute("manage", docPath, outputPath: outputPath,
            revisionIndex: 0, action: "accept");
        Assert.Contains("[0]", result);
        Assert.Contains("accepted", result);
    }

    [Fact]
    public void ManageRevision_Reject_ShouldRejectSpecificRevision()
    {
        var docPath = CreateWordDocument("test_manage_reject.docx");
        var doc = new Document(docPath);
        doc.StartTrackRevisions("Test Author");
        var builder = new DocumentBuilder(doc);
        builder.Writeln("Revision to reject");
        doc.StopTrackRevisions();
        doc.Save(docPath);

        var outputPath = CreateTestFilePath("test_manage_reject_output.docx");
        var result = _tool.Execute("manage", docPath, outputPath: outputPath,
            revisionIndex: 0, action: "reject");
        Assert.Contains("[0]", result);
        Assert.Contains("rejected", result);
    }

    [Fact]
    public void ManageRevision_WithInvalidIndex_ShouldThrowException()
    {
        var docPath = CreateWordDocument("test_manage_invalid_index.docx");
        var doc = new Document(docPath);
        doc.StartTrackRevisions("Test Author");
        var builder = new DocumentBuilder(doc);
        builder.Writeln("Single revision");
        doc.StopTrackRevisions();
        doc.Save(docPath);
        var exception = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("manage", docPath, revisionIndex: 99, action: "accept"));
        Assert.Contains("revisionIndex must be between", exception.Message);
    }

    [Fact]
    public void ManageRevision_WithNoRevisions_ShouldReturnMessage()
    {
        var docPath = CreateWordDocumentWithContent("test_manage_no_rev.docx", "Plain text");
        var result = _tool.Execute("manage", docPath, revisionIndex: 0, action: "accept");
        Assert.Contains("no revisions", result);
    }

    [Fact]
    public void CompareDocuments_ShouldCreateComparisonDocument()
    {
        var originalPath = CreateWordDocumentWithContent("test_compare_original.docx", "Original content");
        var revisedPath = CreateWordDocumentWithContent("test_compare_revised.docx", "Revised content");
        var outputPath = CreateTestFilePath("test_compare_output.docx");
        var result = _tool.Execute("compare", outputPath: outputPath,
            originalPath: originalPath, revisedPath: revisedPath);
        Assert.True(File.Exists(outputPath));
        Assert.Contains("Comparison completed", result);
        Assert.Contains("difference(s) found", result);
    }

    [Fact]
    public void CompareDocuments_WithIgnoreFormatting_ShouldWork()
    {
        var originalPath = CreateWordDocumentWithContent("test_compare_fmt_orig.docx", "Same content");
        var revisedPath = CreateWordDocumentWithContent("test_compare_fmt_rev.docx", "Same content");
        var outputPath = CreateTestFilePath("test_compare_fmt_output.docx");
        var result = _tool.Execute("compare", outputPath: outputPath,
            originalPath: originalPath, revisedPath: revisedPath,
            ignoreFormatting: true, ignoreComments: true);
        Assert.True(File.Exists(outputPath));
        Assert.Contains("Comparison completed", result);
    }

    #endregion

    #region Exception Tests

    [Fact]
    public void Execute_WithUnknownOperation_ShouldThrowArgumentException()
    {
        var docPath = CreateWordDocumentWithContent("test_unknown_op.docx", "Test content");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("unknown_operation", docPath));

        Assert.Contains("Unknown operation", ex.Message);
        Assert.Contains("unknown_operation", ex.Message);
    }

    [Fact]
    public void ManageRevision_WithMissingAction_ShouldThrowArgumentException()
    {
        var docPath = CreateWordDocument("test_missing_action.docx");
        var doc = new Document(docPath);
        doc.StartTrackRevisions("Test Author");
        var builder = new DocumentBuilder(doc);
        builder.Writeln("Revision text");
        doc.StopTrackRevisions();
        doc.Save(docPath);
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("manage", docPath, revisionIndex: 0, action: ""));

        Assert.Contains("action", ex.Message);
    }

    [Fact]
    public void Compare_WithMissingOriginalPath_ShouldThrowArgumentException()
    {
        var revisedPath = CreateWordDocumentWithContent("test_compare_missing.docx", "Content");
        var outputPath = CreateTestFilePath("test_compare_missing_output.docx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("compare", outputPath: outputPath, originalPath: "", revisedPath: revisedPath));

        Assert.Contains("originalPath", ex.Message);
    }

    #endregion

    #region Session ID Tests

    [Fact]
    public void GetRevisions_WithSessionId_ShouldReturnRevisions()
    {
        var docPath = CreateWordDocument("test_session_get_revisions.docx");
        var doc = new Document(docPath);
        doc.StartTrackRevisions("Test Author");
        var builder = new DocumentBuilder(doc);
        builder.Writeln("Session revision text");
        doc.StopTrackRevisions();
        doc.Save(docPath);

        var sessionId = OpenSession(docPath);
        var result = _tool.Execute("get_revisions", sessionId: sessionId);
        Assert.Contains("\"revisions\"", result);
    }

    [Fact]
    public void AcceptAllRevisions_WithSessionId_ShouldAcceptInMemory()
    {
        var docPath = CreateWordDocument("test_session_accept_all.docx");
        var doc = new Document(docPath);
        doc.StartTrackRevisions("Test Author");
        var builder = new DocumentBuilder(doc);
        builder.Writeln("Session revision");
        doc.StopTrackRevisions();
        doc.Save(docPath);

        var sessionId = OpenSession(docPath);
        var result = _tool.Execute("accept_all", sessionId: sessionId);
        Assert.Contains("Accepted", result);

        // Verify in-memory document has no revisions
        var sessionDoc = SessionManager.GetDocument<Document>(sessionId);
        Assert.Equal(0, sessionDoc.Revisions.Count);
    }

    [Fact]
    public void ManageRevision_WithSessionId_ShouldManageInMemory()
    {
        var docPath = CreateWordDocument("test_session_manage.docx");
        var doc = new Document(docPath);
        doc.StartTrackRevisions("Test Author");
        var builder = new DocumentBuilder(doc);
        builder.Writeln("Revision to manage");
        doc.StopTrackRevisions();
        doc.Save(docPath);

        var sessionId = OpenSession(docPath);
        var result = _tool.Execute("manage", sessionId: sessionId, revisionIndex: 0, action: "accept");
        Assert.Contains("[0]", result);
        Assert.Contains("accepted", result);
    }

    [Fact]
    public void Execute_WithInvalidSessionId_ShouldThrowKeyNotFoundException()
    {
        Assert.Throws<KeyNotFoundException>(() =>
            _tool.Execute("get_revisions", sessionId: "invalid_session_id"));
    }

    [Fact]
    public void Execute_WithBothPathAndSessionId_ShouldPreferSessionId()
    {
        var docPath1 = CreateWordDocument("test_path_rev.docx");
        var doc1 = new Document(docPath1);
        doc1.StartTrackRevisions("Author1");
        var builder1 = new DocumentBuilder(doc1);
        builder1.Writeln("Path revision");
        doc1.StopTrackRevisions();
        doc1.Save(docPath1);

        var docPath2 = CreateWordDocument("test_session_rev.docx");
        var doc2 = new Document(docPath2);
        doc2.StartTrackRevisions("Author2");
        var builder2 = new DocumentBuilder(doc2);
        builder2.Writeln("Session revision");
        doc2.StopTrackRevisions();
        doc2.Save(docPath2);

        var sessionId = OpenSession(docPath2);

        // Act - provide both path and sessionId
        var result = _tool.Execute("get_revisions", docPath1, sessionId);

        // Assert - should use sessionId
        Assert.Contains("\"revisions\"", result);
    }

    #endregion
}