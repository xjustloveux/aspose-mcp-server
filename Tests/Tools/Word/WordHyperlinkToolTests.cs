using Aspose.Words;
using Aspose.Words.Fields;
using AsposeMcpServer.Tests.Helpers;
using AsposeMcpServer.Tools.Word;

namespace AsposeMcpServer.Tests.Tools.Word;

public class WordHyperlinkToolTests : WordTestBase
{
    private readonly WordHyperlinkTool _tool;

    public WordHyperlinkToolTests()
    {
        _tool = new WordHyperlinkTool(SessionManager);
    }

    #region General Tests

    [Fact]
    public void AddHyperlink_ShouldAddHyperlink()
    {
        var docPath = CreateWordDocumentWithContent("test_add_hyperlink.docx", "Test content");
        var outputPath = CreateTestFilePath("test_add_hyperlink_output.docx");
        _tool.Execute("add", docPath, outputPath: outputPath,
            text: "Click here", url: "https://example.com", paragraphIndex: 0);
        var doc = new Document(outputPath);
        var hyperlinks = doc.Range.Fields.Where(f => f.Type == FieldType.FieldHyperlink).ToList();
        Assert.True(hyperlinks.Count > 0 || doc.GetText().Contains("Click here"),
            "Document should contain a hyperlink or the hyperlink text");
    }

    [Fact]
    public void GetHyperlinks_ShouldReturnAllHyperlinks()
    {
        var docPath = CreateWordDocumentWithContent("test_get_hyperlinks.docx", "Test content");
        var doc = new Document(docPath);
        var builder = new DocumentBuilder(doc);
        builder.InsertHyperlink("Test Link", "https://test.com", false);
        doc.Save(docPath);
        var result = _tool.Execute("get", docPath);
        Assert.NotNull(result);
        Assert.NotEmpty(result);
    }

    [Fact]
    public void EditHyperlink_ShouldEditHyperlink()
    {
        var docPath = CreateWordDocumentWithContent("test_edit_hyperlink.docx", "Test content");
        var doc = new Document(docPath);
        var builder = new DocumentBuilder(doc);
        builder.InsertHyperlink("Original Link", "https://original.com", false);
        doc.Save(docPath);

        var outputPath = CreateTestFilePath("test_edit_hyperlink_output.docx");
        _tool.Execute("edit", docPath, outputPath: outputPath,
            hyperlinkIndex: 0, url: "https://updated.com");
        Assert.True(File.Exists(outputPath), "Output document should be created");
        var resultDoc = new Document(outputPath);
        Assert.NotNull(resultDoc);
    }

    [Fact]
    public void DeleteHyperlink_ShouldDeleteHyperlink()
    {
        var docPath = CreateWordDocumentWithContent("test_delete_hyperlink.docx", "Test content");
        var doc = new Document(docPath);
        var builder = new DocumentBuilder(doc);
        builder.InsertHyperlink("Link to Delete", "https://delete.com", false);
        doc.Save(docPath);

        var outputPath = CreateTestFilePath("test_delete_hyperlink_output.docx");
        _tool.Execute("delete", docPath, outputPath: outputPath, hyperlinkIndex: 0);
        Assert.True(File.Exists(outputPath), "Output document should be created");
        var resultDoc = new Document(outputPath);
        var hyperlinks = resultDoc.Range.Fields.Where(f => f.Type == FieldType.FieldHyperlink).ToList();
        Assert.Empty(hyperlinks);
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
    public void AddHyperlink_WithMissingUrl_ShouldThrowArgumentException()
    {
        var docPath = CreateWordDocumentWithContent("test_add_missing_url.docx", "Test content");
        var outputPath = CreateTestFilePath("test_add_missing_url_output.docx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("add", docPath, outputPath: outputPath, text: "Link Text", url: null));

        Assert.Contains("Either 'url' or 'subAddress' must be provided", ex.Message);
    }

    [Fact]
    public void EditHyperlink_WithInvalidIndex_ShouldThrowArgumentException()
    {
        var docPath = CreateWordDocumentWithContent("test_edit_invalid_index.docx", "Test content");
        var doc = new Document(docPath);
        var builder = new DocumentBuilder(doc);
        builder.InsertHyperlink("Single Link", "https://example.com", false);
        doc.Save(docPath);

        var outputPath = CreateTestFilePath("test_edit_invalid_index_output.docx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("edit", docPath, outputPath: outputPath, hyperlinkIndex: 999, url: "https://new.com"));

        Assert.Contains("out of range", ex.Message);
    }

    [Fact]
    public void DeleteHyperlink_WithInvalidIndex_ShouldThrowArgumentException()
    {
        var docPath = CreateWordDocumentWithContent("test_delete_invalid_index.docx", "Test content");
        var outputPath = CreateTestFilePath("test_delete_invalid_index_output.docx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("delete", docPath, outputPath: outputPath, hyperlinkIndex: 0));

        Assert.Contains("out of range", ex.Message);
    }

    #endregion

    #region Session ID Tests

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
        Assert.Contains("session.com", result, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void AddHyperlink_WithSessionId_ShouldAddHyperlinkInMemory()
    {
        var docPath = CreateWordDocumentWithContent("test_session_add_hl.docx", "Test content for hyperlink");
        var sessionId = OpenSession(docPath);
        var result = _tool.Execute("add", sessionId: sessionId,
            text: "Session Hyperlink", url: "https://session-link.com", paragraphIndex: 0);
        Assert.Contains("Session Hyperlink", result);

        // Verify in-memory document has the hyperlink
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

        // Assert - verify in-memory change
        var sessionDoc = SessionManager.GetDocument<Document>(sessionId);
        var hyperlinks = sessionDoc.Range.Fields.Where(f => f.Type == FieldType.FieldHyperlink).ToList();
        Assert.NotEmpty(hyperlinks);
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

        // Assert - verify in-memory deletion
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

        // Act - provide both path and sessionId
        var result = _tool.Execute("get", docPath1, sessionId);

        // Assert - should use sessionId, returning Session Link not Path Link
        Assert.Contains("session-unique.com", result);
        Assert.DoesNotContain("path-unique.com", result);
    }

    #endregion
}