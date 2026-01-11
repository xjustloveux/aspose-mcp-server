using Aspose.Words;
using AsposeMcpServer.Tests.Helpers;
using AsposeMcpServer.Tools.Word;

namespace AsposeMcpServer.Tests.Tools.Word;

/// <summary>
///     Integration tests for WordRevisionTool.
///     Focuses on session management, file I/O, and operation routing.
///     Detailed parameter validation and business logic tests are in Handler tests.
/// </summary>
public class WordRevisionToolTests : WordTestBase
{
    private readonly WordRevisionTool _tool;

    public WordRevisionToolTests()
    {
        _tool = new WordRevisionTool(SessionManager);
    }

    #region File I/O Smoke Tests

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
        Assert.Contains("\"revisions\"", result);
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
        Assert.StartsWith("Accepted", result);
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
        Assert.StartsWith("Rejected", result);
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
    public void CompareDocuments_ShouldCreateComparisonDocument()
    {
        var originalPath = CreateWordDocumentWithContent("test_compare_original.docx", "Original content");
        var revisedPath = CreateWordDocumentWithContent("test_compare_revised.docx", "Revised content");
        var outputPath = CreateTestFilePath("test_compare_output.docx");
        var result = _tool.Execute("compare", outputPath: outputPath,
            originalPath: originalPath, revisedPath: revisedPath);
        Assert.True(File.Exists(outputPath));
        Assert.StartsWith("Comparison completed", result);
        Assert.Contains("difference(s) found", result);
    }

    #endregion

    #region Operation Routing

    [Theory]
    [InlineData("GET_REVISIONS")]
    [InlineData("Get_Revisions")]
    [InlineData("get_revisions")]
    public void Operation_ShouldBeCaseInsensitive(string operation)
    {
        var docPath = CreateWordDocumentWithContent($"test_case_{operation.Replace("_", "")}.docx", "Content");
        var result = _tool.Execute(operation, docPath);
        Assert.Contains("\"revisions\"", result);
    }

    [Fact]
    public void Execute_WithUnknownOperation_ShouldThrowArgumentException()
    {
        var docPath = CreateWordDocumentWithContent("test_unknown_op.docx", "Test content");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("unknown_operation", docPath));
        Assert.Contains("Unknown operation", ex.Message);
    }

    [Fact]
    public void Execute_WithNoPathOrSessionId_ShouldThrowException()
    {
        Assert.ThrowsAny<Exception>(() =>
            _tool.Execute("get_revisions"));
    }

    #endregion

    #region Session Management

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
        Assert.StartsWith("Accepted", result);

        var sessionDoc = SessionManager.GetDocument<Document>(sessionId);
        Assert.Equal(0, sessionDoc.Revisions.Count);
    }

    [Fact]
    public void RejectAllRevisions_WithSessionId_ShouldRejectInMemory()
    {
        var docPath = CreateWordDocument("test_session_reject_all.docx");
        var doc = new Document(docPath);
        doc.StartTrackRevisions("Test Author");
        var builder = new DocumentBuilder(doc);
        builder.Writeln("Session revision to reject");
        doc.StopTrackRevisions();
        doc.Save(docPath);

        var sessionId = OpenSession(docPath);
        var result = _tool.Execute("reject_all", sessionId: sessionId);
        Assert.StartsWith("Rejected", result);

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
        var result = _tool.Execute("get_revisions", docPath1, sessionId);
        Assert.Contains("\"revisions\"", result);
    }

    #endregion
}
