using Aspose.Words;
using Aspose.Words.Fields;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Results.Word.Hyperlink;
using AsposeMcpServer.Tests.Infrastructure;
using AsposeMcpServer.Tools.Word;

namespace AsposeMcpServer.Tests.Tools.Word;

/// <summary>
///     Integration tests for WordHyperlinkTool.
///     Focuses on session management, file I/O, and operation routing.
///     Detailed parameter validation and business logic tests are in Handler tests.
/// </summary>
public class WordHyperlinkToolTests : WordTestBase
{
    private readonly WordHyperlinkTool _tool;

    public WordHyperlinkToolTests()
    {
        _tool = new WordHyperlinkTool(SessionManager);
    }

    #region File I/O Smoke Tests

    [Fact]
    public void AddHyperlink_ShouldAddHyperlinkAndPersistToFile()
    {
        var docPath = CreateWordDocumentWithContent("test_add_hyperlink.docx", "Test content");
        var outputPath = CreateTestFilePath("test_add_hyperlink_output.docx");
        _tool.Execute("add", docPath, outputPath: outputPath,
            text: "Click here", url: "https://example.com", paragraphIndex: 0);
        var doc = new Document(outputPath);
        var hyperlinks = doc.Range.Fields.Where(f => f.Type == FieldType.FieldHyperlink).ToList();
        Assert.True(hyperlinks.Count > 0 || doc.GetText().Contains("Click here"));
    }

    [Fact]
    public void GetHyperlinks_ShouldReturnHyperlinksFromFile()
    {
        var docPath = CreateWordDocumentWithContent("test_get_hyperlinks.docx", "Test content");
        var doc = new Document(docPath);
        var builder = new DocumentBuilder(doc);
        builder.InsertHyperlink("Test Link", "https://test.com", false);
        doc.Save(docPath);
        var result = _tool.Execute("get", docPath);
        var data = GetResultData<GetHyperlinksResult>(result);
        Assert.True(data.Count > 0);
        Assert.Contains(data.Hyperlinks, h => h.Address.Contains("test.com"));
    }

    [Fact]
    public void EditHyperlink_ShouldEditHyperlinkAndPersistToFile()
    {
        var docPath = CreateWordDocumentWithContent("test_edit_hyperlink.docx", "Test content");
        var doc = new Document(docPath);
        var builder = new DocumentBuilder(doc);
        builder.InsertHyperlink("Original Link", "https://original.com", false);
        doc.Save(docPath);

        var outputPath = CreateTestFilePath("test_edit_hyperlink_output.docx");
        _tool.Execute("edit", docPath, outputPath: outputPath,
            hyperlinkIndex: 0, url: "https://updated.com");
        var resultDoc = new Document(outputPath);
        var hyperlinks = resultDoc.Range.Fields.OfType<FieldHyperlink>().ToList();
        Assert.Equal("https://updated.com", hyperlinks[0].Address);
    }

    [Fact]
    public void DeleteHyperlink_ShouldDeleteHyperlinkAndPersistToFile()
    {
        var docPath = CreateWordDocumentWithContent("test_delete_hyperlink.docx", "Test content");
        var doc = new Document(docPath);
        var builder = new DocumentBuilder(doc);
        builder.InsertHyperlink("Link to Delete", "https://delete.com", false);
        doc.Save(docPath);

        var outputPath = CreateTestFilePath("test_delete_hyperlink_output.docx");
        _tool.Execute("delete", docPath, outputPath: outputPath, hyperlinkIndex: 0);
        var resultDoc = new Document(outputPath);
        var hyperlinks = resultDoc.Range.Fields.Where(f => f.Type == FieldType.FieldHyperlink).ToList();
        Assert.Empty(hyperlinks);
    }

    #endregion

    #region Operation Routing

    [Theory]
    [InlineData("ADD")]
    [InlineData("Add")]
    [InlineData("add")]
    public void Operation_ShouldBeCaseInsensitive(string operation)
    {
        var docPath = CreateWordDocumentWithContent($"test_case_{operation}.docx", "Test content");
        var outputPath = CreateTestFilePath($"test_case_{operation}_output.docx");
        var result = _tool.Execute(operation, docPath, outputPath: outputPath,
            text: "Link", url: "https://example.com", paragraphIndex: 0);
        var data = GetResultData<SuccessResult>(result);
        Assert.StartsWith("Hyperlink added", data.Message);
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
        Assert.ThrowsAny<Exception>(() => _tool.Execute("get"));
    }

    #endregion

    #region Session Management

    [Fact]
    public void GetHyperlinks_WithSessionId_ShouldReturnHyperlinks()
    {
        var docPath = CreateWordDocumentWithContent("test_session_get_hl.docx", "Test content");
        var doc = new Document(docPath);
        var builder = new DocumentBuilder(doc);
        builder.InsertHyperlink("Session Link", "https://session.com", false);
        doc.Save(docPath);

        var sessionId = OpenSession(docPath);
        var result = _tool.Execute("get", sessionId: sessionId);
        var data = GetResultData<GetHyperlinksResult>(result);
        Assert.True(data.Count > 0);
        Assert.Contains(data.Hyperlinks, h => h.Address.Contains("session.com", StringComparison.OrdinalIgnoreCase));
    }

    [Fact]
    public void AddHyperlink_WithSessionId_ShouldAddHyperlinkInMemory()
    {
        var docPath = CreateWordDocumentWithContent("test_session_add_hl.docx", "Test content for hyperlink");
        var sessionId = OpenSession(docPath);
        var result = _tool.Execute("add", sessionId: sessionId,
            text: "Session Hyperlink", url: "https://session-link.com", paragraphIndex: 0);
        var data = GetResultData<SuccessResult>(result);
        Assert.Contains("Session Hyperlink", data.Message);

        var doc = SessionManager.GetDocument<Document>(sessionId);
        var hyperlinks = doc.Range.Fields.Where(f => f.Type == FieldType.FieldHyperlink).ToList();
        Assert.True(hyperlinks.Count > 0 || doc.GetText().Contains("Session Hyperlink"));
    }

    [Fact]
    public void EditHyperlink_WithSessionId_ShouldEditHyperlinkInMemory()
    {
        var docPath = CreateWordDocumentWithContent("test_session_edit_hl.docx", "Test content");
        var doc = new Document(docPath);
        var builder = new DocumentBuilder(doc);
        builder.InsertHyperlink("Original Session Link", "https://original-session.com", false);
        doc.Save(docPath);

        var sessionId = OpenSession(docPath);
        _tool.Execute("edit", sessionId: sessionId, hyperlinkIndex: 0, url: "https://updated-session.com");

        var sessionDoc = SessionManager.GetDocument<Document>(sessionId);
        var hyperlinks = sessionDoc.Range.Fields.Where(f => f.Type == FieldType.FieldHyperlink).ToList();
        var hyperlinkField = (FieldHyperlink)hyperlinks[0];
        Assert.Equal("https://updated-session.com", hyperlinkField.Address);
    }

    [Fact]
    public void DeleteHyperlink_WithSessionId_ShouldDeleteHyperlinkInMemory()
    {
        var docPath = CreateWordDocumentWithContent("test_session_delete_hl.docx", "Test content");
        var doc = new Document(docPath);
        var builder = new DocumentBuilder(doc);
        builder.InsertHyperlink("Link To Delete Session", "https://delete-session.com", false);
        doc.Save(docPath);

        var sessionId = OpenSession(docPath);
        _tool.Execute("delete", sessionId: sessionId, hyperlinkIndex: 0);

        var sessionDoc = SessionManager.GetDocument<Document>(sessionId);
        var hyperlinks = sessionDoc.Range.Fields.Where(f => f.Type == FieldType.FieldHyperlink).ToList();
        Assert.Empty(hyperlinks);
    }

    [Fact]
    public void Execute_WithInvalidSessionId_ShouldThrowKeyNotFoundException()
    {
        Assert.Throws<KeyNotFoundException>(() =>
            _tool.Execute("get", sessionId: "invalid_session_id"));
    }

    [Fact]
    public void Execute_WithBothPathAndSessionId_ShouldPreferSessionId()
    {
        var docPath1 = CreateWordDocumentWithContent("test_path_hl.docx", "Path content");
        var doc1 = new Document(docPath1);
        var builder1 = new DocumentBuilder(doc1);
        builder1.InsertHyperlink("Path Link Unique", "https://path-unique.com", false);
        doc1.Save(docPath1);

        var docPath2 = CreateWordDocumentWithContent("test_session_hl.docx", "Session content");
        var doc2 = new Document(docPath2);
        var builder2 = new DocumentBuilder(doc2);
        builder2.InsertHyperlink("Session Link Unique", "https://session-unique.com", false);
        doc2.Save(docPath2);

        var sessionId = OpenSession(docPath2);
        var result = _tool.Execute("get", docPath1, sessionId);
        var data = GetResultData<GetHyperlinksResult>(result);

        Assert.Contains(data.Hyperlinks, h => h.Address.Contains("session-unique.com"));
        Assert.DoesNotContain(data.Hyperlinks, h => h.Address.Contains("path-unique.com"));
    }

    #endregion
}
