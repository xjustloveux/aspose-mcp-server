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

    #region General

    [Fact]
    public void AddHyperlink_ShouldAddHyperlink()
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
        Assert.Contains("test.com", result);
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
        Assert.True(File.Exists(outputPath));
        var resultDoc = new Document(outputPath);
        var hyperlinks = resultDoc.Range.Fields.OfType<FieldHyperlink>().ToList();
        Assert.NotEmpty(hyperlinks);
        Assert.Equal("https://updated.com", hyperlinks[0].Address);
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
        Assert.True(File.Exists(outputPath));
        var resultDoc = new Document(outputPath);
        var hyperlinks = resultDoc.Range.Fields.Where(f => f.Type == FieldType.FieldHyperlink).ToList();
        Assert.Empty(hyperlinks);
    }

    [Theory]
    [InlineData("ADD")]
    [InlineData("Add")]
    [InlineData("add")]
    public void Operation_ShouldBeCaseInsensitive_Add(string operation)
    {
        var docPath = CreateWordDocumentWithContent($"test_case_{operation}.docx", "Test content");
        var outputPath = CreateTestFilePath($"test_case_{operation}_output.docx");
        var result = _tool.Execute(operation, docPath, outputPath: outputPath,
            text: "Link", url: "https://example.com", paragraphIndex: 0);
        Assert.StartsWith("Hyperlink added", result);
    }

    [Theory]
    [InlineData("GET")]
    [InlineData("Get")]
    [InlineData("get")]
    public void Operation_ShouldBeCaseInsensitive_Get(string operation)
    {
        var docPath = CreateWordDocumentWithContent($"test_case_get_{operation}.docx", "Test content");
        var doc = new Document(docPath);
        var builder = new DocumentBuilder(doc);
        builder.InsertHyperlink("Link", "https://test.com", false);
        doc.Save(docPath);
        var result = _tool.Execute(operation, docPath);
        Assert.Contains("hyperlinks", result.ToLower());
    }

    [Theory]
    [InlineData("EDIT")]
    [InlineData("Edit")]
    [InlineData("edit")]
    public void Operation_ShouldBeCaseInsensitive_Edit(string operation)
    {
        var docPath = CreateWordDocumentWithContent($"test_case_edit_{operation}.docx", "Test content");
        var doc = new Document(docPath);
        var builder = new DocumentBuilder(doc);
        builder.InsertHyperlink("Link", "https://original.com", false);
        doc.Save(docPath);
        var outputPath = CreateTestFilePath($"test_case_edit_{operation}_output.docx");
        var result = _tool.Execute(operation, docPath, outputPath: outputPath,
            hyperlinkIndex: 0, url: "https://new.com");
        Assert.StartsWith("Hyperlink #0 edited", result);
    }

    [Theory]
    [InlineData("DELETE")]
    [InlineData("Delete")]
    [InlineData("delete")]
    public void Operation_ShouldBeCaseInsensitive_Delete(string operation)
    {
        var docPath = CreateWordDocumentWithContent($"test_case_delete_{operation}.docx", "Test content");
        var doc = new Document(docPath);
        var builder = new DocumentBuilder(doc);
        builder.InsertHyperlink("Link", "https://delete.com", false);
        doc.Save(docPath);
        var outputPath = CreateTestFilePath($"test_case_delete_{operation}_output.docx");
        var result = _tool.Execute(operation, docPath, outputPath: outputPath, hyperlinkIndex: 0);
        Assert.StartsWith("Hyperlink #0 deleted", result);
    }

    [Fact]
    public void AddHyperlink_WithSubAddress_ShouldAddInternalLink()
    {
        var docPath = CreateWordDocumentWithContent("test_add_subaddress.docx", "Test content");
        var outputPath = CreateTestFilePath("test_add_subaddress_output.docx");
        var result = _tool.Execute("add", docPath, outputPath: outputPath,
            text: "Go to bookmark", subAddress: "_Toc123456", paragraphIndex: 0);
        Assert.StartsWith("Hyperlink added", result);
        Assert.Contains("SubAddress", result);
    }

    [Fact]
    public void AddHyperlink_WithTooltip_ShouldSetTooltip()
    {
        var docPath = CreateWordDocumentWithContent("test_add_tooltip.docx", "Test content");
        var outputPath = CreateTestFilePath("test_add_tooltip_output.docx");
        var result = _tool.Execute("add", docPath, outputPath: outputPath,
            text: "Hover me", url: "https://example.com", tooltip: "This is a tooltip", paragraphIndex: 0);
        Assert.StartsWith("Hyperlink added", result);
        Assert.Contains("Tooltip:", result);
    }

    [Fact]
    public void AddHyperlink_AtDocumentStart_ShouldInsertAtBeginning()
    {
        var docPath = CreateWordDocumentWithContent("test_add_start.docx", "Existing content");
        var outputPath = CreateTestFilePath("test_add_start_output.docx");
        var result = _tool.Execute("add", docPath, outputPath: outputPath,
            text: "Start Link", url: "https://start.com", paragraphIndex: -1);
        Assert.StartsWith("Hyperlink added", result);
        Assert.Contains("beginning of document", result);
    }

    [Fact]
    public void AddHyperlink_AtDocumentEnd_ShouldInsertAtEnd()
    {
        var docPath = CreateWordDocumentWithContent("test_add_end.docx", "Existing content");
        var outputPath = CreateTestFilePath("test_add_end_output.docx");
        var result = _tool.Execute("add", docPath, outputPath: outputPath,
            text: "End Link", url: "https://end.com");
        Assert.StartsWith("Hyperlink added", result);
        Assert.Contains("end of document", result);
    }

    [Fact]
    public void EditHyperlink_WithDisplayText_ShouldUpdateDisplayText()
    {
        var docPath = CreateWordDocumentWithContent("test_edit_displaytext.docx", "Test content");
        var doc = new Document(docPath);
        var builder = new DocumentBuilder(doc);
        builder.InsertHyperlink("Original Text", "https://example.com", false);
        doc.Save(docPath);

        var outputPath = CreateTestFilePath("test_edit_displaytext_output.docx");
        var result = _tool.Execute("edit", docPath, outputPath: outputPath,
            hyperlinkIndex: 0, displayText: "New Display Text");
        Assert.StartsWith("Hyperlink #0 edited", result);
        Assert.Contains("Display text:", result);
    }

    [Fact]
    public void EditHyperlink_WithTooltip_ShouldUpdateTooltip()
    {
        var docPath = CreateWordDocumentWithContent("test_edit_tooltip.docx", "Test content");
        var doc = new Document(docPath);
        var builder = new DocumentBuilder(doc);
        builder.InsertHyperlink("Link", "https://example.com", false);
        doc.Save(docPath);

        var outputPath = CreateTestFilePath("test_edit_tooltip_output.docx");
        var result = _tool.Execute("edit", docPath, outputPath: outputPath,
            hyperlinkIndex: 0, tooltip: "New tooltip text");
        Assert.StartsWith("Hyperlink #0 edited", result);
        Assert.Contains("Tooltip:", result);
    }

    [Fact]
    public void DeleteHyperlink_WithKeepText_ShouldUnlinkButKeepText()
    {
        var docPath = CreateWordDocumentWithContent("test_delete_keeptext.docx", "Test content");
        var doc = new Document(docPath);
        var builder = new DocumentBuilder(doc);
        builder.InsertHyperlink("Keep This Text", "https://delete.com", false);
        doc.Save(docPath);

        var outputPath = CreateTestFilePath("test_delete_keeptext_output.docx");
        var result = _tool.Execute("delete", docPath, outputPath: outputPath,
            hyperlinkIndex: 0, keepText: true);
        Assert.StartsWith("Hyperlink #0 deleted", result);
        Assert.Contains("unlinked", result);

        var resultDoc = new Document(outputPath);
        Assert.Contains("Keep This Text", resultDoc.GetText());
        var hyperlinks = resultDoc.Range.Fields.Where(f => f.Type == FieldType.FieldHyperlink).ToList();
        Assert.Empty(hyperlinks);
    }

    [Fact]
    public void GetHyperlinks_WithNoHyperlinks_ShouldReturnEmptyResult()
    {
        var docPath = CreateWordDocumentWithContent("test_get_no_hyperlinks.docx", "Text without hyperlinks");
        var result = _tool.Execute("get", docPath);
        Assert.Contains("\"count\":0", result);
        Assert.Contains("No hyperlinks found", result);
    }

    [Fact]
    public void GetHyperlinks_WithMultipleHyperlinks_ShouldReturnAll()
    {
        var docPath = CreateWordDocumentWithContent("test_get_multiple.docx", "Test content");
        var doc = new Document(docPath);
        var builder = new DocumentBuilder(doc);
        builder.InsertHyperlink("Link 1", "https://link1.com", false);
        builder.Writeln();
        builder.InsertHyperlink("Link 2", "https://link2.com", false);
        builder.Writeln();
        builder.InsertHyperlink("Link 3", "https://link3.com", false);
        doc.Save(docPath);

        var result = _tool.Execute("get", docPath);
        Assert.Contains("\"count\": 3", result);
        Assert.Contains("link1.com", result);
        Assert.Contains("link2.com", result);
        Assert.Contains("link3.com", result);
    }

    [Fact]
    public void AddHyperlink_WithMailtoUrl_ShouldAddEmailLink()
    {
        var docPath = CreateWordDocumentWithContent("test_add_mailto.docx", "Test content");
        var outputPath = CreateTestFilePath("test_add_mailto_output.docx");
        var result = _tool.Execute("add", docPath, outputPath: outputPath,
            text: "Email us", url: "mailto:test@example.com", paragraphIndex: 0);
        Assert.StartsWith("Hyperlink added", result);
        Assert.Contains("mailto:", result);
    }

    [Fact]
    public void AddHyperlink_WithFileUrl_ShouldAddFileLink()
    {
        var docPath = CreateWordDocumentWithContent("test_add_file.docx", "Test content");
        var outputPath = CreateTestFilePath("test_add_file_output.docx");
        var result = _tool.Execute("add", docPath, outputPath: outputPath,
            text: "Open file", url: "file:///C:/test.txt", paragraphIndex: 0);
        Assert.StartsWith("Hyperlink added", result);
        Assert.Contains("file://", result);
    }

    #endregion

    #region Exception

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
    public void AddHyperlink_WithMissingUrlAndSubAddress_ShouldThrowArgumentException()
    {
        var docPath = CreateWordDocumentWithContent("test_add_missing_url.docx", "Test content");
        var outputPath = CreateTestFilePath("test_add_missing_url_output.docx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("add", docPath, outputPath: outputPath, text: "Link Text", url: null, subAddress: null));

        Assert.Contains("Either 'url' or 'subAddress' must be provided", ex.Message);
    }

    [Fact]
    public void AddHyperlink_WithInvalidUrl_ShouldThrowArgumentException()
    {
        var docPath = CreateWordDocumentWithContent("test_add_invalid_url.docx", "Test content");
        var outputPath = CreateTestFilePath("test_add_invalid_url_output.docx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("add", docPath, outputPath: outputPath, text: "Link", url: "invalid-url"));

        Assert.Contains("Invalid URL format", ex.Message);
    }

    [Fact]
    public void AddHyperlink_WithInvalidParagraphIndex_ShouldThrowArgumentException()
    {
        var docPath = CreateWordDocumentWithContent("test_add_invalid_para.docx", "Single paragraph");
        var outputPath = CreateTestFilePath("test_add_invalid_para_output.docx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("add", docPath, outputPath: outputPath, text: "Link",
                url: "https://example.com", paragraphIndex: 999));

        Assert.Contains("out of range", ex.Message);
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
    public void EditHyperlink_WithNoHyperlinks_ShouldThrowArgumentException()
    {
        var docPath = CreateWordDocumentWithContent("test_edit_no_hyperlinks.docx", "No hyperlinks");
        var outputPath = CreateTestFilePath("test_edit_no_hyperlinks_output.docx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("edit", docPath, outputPath: outputPath, hyperlinkIndex: 0, url: "https://new.com"));

        Assert.Contains("out of range", ex.Message);
        Assert.Contains("has no hyperlinks", ex.Message);
    }

    [Fact]
    public void EditHyperlink_WithInvalidUrl_ShouldThrowArgumentException()
    {
        var docPath = CreateWordDocumentWithContent("test_edit_invalid_url.docx", "Test content");
        var doc = new Document(docPath);
        var builder = new DocumentBuilder(doc);
        builder.InsertHyperlink("Link", "https://example.com", false);
        doc.Save(docPath);

        var outputPath = CreateTestFilePath("test_edit_invalid_url_output.docx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("edit", docPath, outputPath: outputPath, hyperlinkIndex: 0, url: "invalid-url"));

        Assert.Contains("Invalid URL format", ex.Message);
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

    [Fact]
    public void DeleteHyperlink_WithNoHyperlinks_ShouldThrowArgumentException()
    {
        var docPath = CreateWordDocumentWithContent("test_delete_no_hyperlinks.docx", "No hyperlinks");
        var outputPath = CreateTestFilePath("test_delete_no_hyperlinks_output.docx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("delete", docPath, outputPath: outputPath, hyperlinkIndex: 0));

        Assert.Contains("out of range", ex.Message);
        Assert.Contains("has no hyperlinks", ex.Message);
    }

    #endregion

    #region Session

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

        Assert.Contains("session-unique.com", result);
        Assert.DoesNotContain("path-unique.com", result);
    }

    [Fact]
    public void AddHyperlink_WithSessionIdAndTooltip_ShouldAddWithTooltip()
    {
        var docPath = CreateWordDocumentWithContent("test_session_tooltip.docx", "Test content");
        var sessionId = OpenSession(docPath);
        var result = _tool.Execute("add", sessionId: sessionId,
            text: "Tooltip Link", url: "https://tooltip.com", tooltip: "Session tooltip", paragraphIndex: 0);
        Assert.Contains("Tooltip:", result);

        var doc = SessionManager.GetDocument<Document>(sessionId);
        var hyperlinks = doc.Range.Fields.OfType<FieldHyperlink>().ToList();
        Assert.NotEmpty(hyperlinks);
    }

    [Fact]
    public void DeleteHyperlink_WithSessionIdAndKeepText_ShouldUnlinkInMemory()
    {
        var docPath = CreateWordDocumentWithContent("test_session_keeptext.docx", "Test content");
        var doc = new Document(docPath);
        var builder = new DocumentBuilder(doc);
        builder.InsertHyperlink("Keep Session Text", "https://session-keep.com", false);
        doc.Save(docPath);

        var sessionId = OpenSession(docPath);
        _tool.Execute("delete", sessionId: sessionId, hyperlinkIndex: 0, keepText: true);

        var sessionDoc = SessionManager.GetDocument<Document>(sessionId);
        Assert.Contains("Keep Session Text", sessionDoc.GetText());
        var hyperlinks = sessionDoc.Range.Fields.Where(f => f.Type == FieldType.FieldHyperlink).ToList();
        Assert.Empty(hyperlinks);
    }

    #endregion
}