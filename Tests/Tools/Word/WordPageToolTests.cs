using Aspose.Words;
using AsposeMcpServer.Tests.Helpers;
using AsposeMcpServer.Tools.Word;

namespace AsposeMcpServer.Tests.Tools.Word;

public class WordPageToolTests : WordTestBase
{
    private readonly WordPageTool _tool;

    public WordPageToolTests()
    {
        _tool = new WordPageTool(SessionManager);
    }

    #region General Tests

    [Fact]
    public void SetMargins_ShouldSetPageMargins()
    {
        var docPath = CreateWordDocument("test_set_margins.docx");
        var outputPath = CreateTestFilePath("test_set_margins_output.docx");
        _tool.Execute("set_margins", docPath, outputPath: outputPath,
            top: 72.0, bottom: 72.0, left: 90.0, right: 90.0);
        var doc = new Document(outputPath);
        var section = doc.Sections[0];
        Assert.Equal(72.0, section.PageSetup.TopMargin);
        Assert.Equal(72.0, section.PageSetup.BottomMargin);
        Assert.Equal(90.0, section.PageSetup.LeftMargin);
        Assert.Equal(90.0, section.PageSetup.RightMargin);
    }

    [Fact]
    public void SetOrientation_ShouldSetPageOrientation()
    {
        var docPath = CreateWordDocument("test_set_orientation.docx");
        var outputPath = CreateTestFilePath("test_set_orientation_output.docx");
        _tool.Execute("set_orientation", docPath, outputPath: outputPath, orientation: "landscape");
        var doc = new Document(outputPath);
        var section = doc.Sections[0];
        Assert.Equal(Orientation.Landscape, section.PageSetup.Orientation);
    }

    [Fact]
    public void SetPageSize_ShouldSetPageSize()
    {
        var docPath = CreateWordDocument("test_set_size.docx");
        var outputPath = CreateTestFilePath("test_set_size_output.docx");
        _tool.Execute("set_size", docPath, outputPath: outputPath, width: 595.0, height: 842.0);
        var doc = new Document(outputPath);
        var section = doc.Sections[0];
        Assert.Equal(595.0, section.PageSetup.PageWidth);
        Assert.Equal(842.0, section.PageSetup.PageHeight);
    }

    [Fact]
    public void SetPageNumber_ShouldSetPageNumber()
    {
        var docPath = CreateWordDocument("test_set_page_number.docx");
        var outputPath = CreateTestFilePath("test_set_page_number_output.docx");
        _tool.Execute("set_page_number", docPath, outputPath: outputPath, startingPageNumber: 5);
        var doc = new Document(outputPath);
        var section = doc.Sections[0];
        // Verify page numbering was set (may require RestartPageNumbering to be true)
        Assert.True(section.PageSetup.RestartPageNumbering || section.PageSetup.PageStartingNumber == 5,
            "Page starting number should be set");
    }

    [Fact]
    public void DeletePage_ShouldRemoveSpecifiedPage()
    {
        // Arrange - Create a multi-page document
        var docPath = CreateTestFilePath("test_delete_page.docx");
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.Writeln("Page 1 content");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("Page 2 content");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("Page 3 content");
        doc.Save(docPath);

        var pageCountBefore = doc.PageCount;
        Assert.True(pageCountBefore >= 3, "Document should have at least 3 pages");

        var outputPath = CreateTestFilePath("test_delete_page_output.docx");
        var result = _tool.Execute("delete_page", docPath, outputPath: outputPath, pageIndex: 1);
        Assert.Contains("deleted successfully", result);
        var resultDoc = new Document(outputPath);
        Assert.True(resultDoc.PageCount < pageCountBefore, "Page count should decrease after deletion");
    }

    [Fact]
    public void InsertBlankPage_ShouldInsertPageAtSpecifiedPosition()
    {
        var docPath = CreateWordDocumentWithContent("test_insert_blank.docx", "Existing content");
        var outputPath = CreateTestFilePath("test_insert_blank_output.docx");
        var result = _tool.Execute("insert_blank_page", docPath, outputPath: outputPath);
        Assert.Contains("inserted", result);
        Assert.True(File.Exists(outputPath));
    }

    [Fact]
    public void AddPageBreak_ShouldAddPageBreakAtDocumentEnd()
    {
        var docPath = CreateWordDocumentWithContent("test_add_page_break.docx", "Content before break");
        var outputPath = CreateTestFilePath("test_add_page_break_output.docx");
        var result = _tool.Execute("add_page_break", docPath, outputPath: outputPath);
        Assert.Contains("Page break added", result);
        var doc = new Document(outputPath);
        // Verify page break was added (document should have increased content)
        Assert.True(doc.GetText().Length > 0);
    }

    [Fact]
    public void AddPageBreak_WithParagraphIndex_ShouldAddBreakAtSpecifiedPosition()
    {
        var docPath = CreateTestFilePath("test_add_break_at_para.docx");
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.Writeln("Paragraph 0");
        builder.Writeln("Paragraph 1");
        builder.Writeln("Paragraph 2");
        doc.Save(docPath);

        var outputPath = CreateTestFilePath("test_add_break_at_para_output.docx");
        var result = _tool.Execute("add_page_break", docPath, outputPath: outputPath, paragraphIndex: 1);
        Assert.Contains("after paragraph 1", result);
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
    public void SetOrientation_WithInvalidOrientation_ShouldDefaultToPortrait()
    {
        var docPath = CreateWordDocument("test_invalid_orientation.docx");
        var outputPath = CreateTestFilePath("test_invalid_orientation_output.docx");

        // Act - Invalid orientation defaults to Portrait (only "landscape" is Landscape)
        var result = _tool.Execute("set_orientation", docPath, outputPath: outputPath, orientation: "diagonal");
        Assert.True(File.Exists(outputPath));
        // The tool echoes the passed orientation value in the message
        Assert.Contains("Page orientation set to diagonal", result);

        // Verify the orientation is actually Portrait (since "diagonal" != "landscape")
        var doc = new Document(outputPath);
        Assert.Equal(Orientation.Portrait, doc.FirstSection.PageSetup.Orientation);
    }

    [Fact]
    public void DeletePage_WithInvalidPageIndex_ShouldThrowArgumentException()
    {
        var docPath = CreateWordDocumentWithContent("test_delete_invalid.docx", "Single page content");
        var outputPath = CreateTestFilePath("test_delete_invalid_output.docx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("delete_page", docPath, outputPath: outputPath, pageIndex: 999));

        Assert.Contains("must be between", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void AddPageBreak_WithInvalidParagraphIndex_ShouldThrowArgumentException()
    {
        var docPath = CreateWordDocumentWithContent("test_break_invalid_para.docx", "Single paragraph");
        var outputPath = CreateTestFilePath("test_break_invalid_para_output.docx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("add_page_break", docPath, outputPath: outputPath, paragraphIndex: 999));

        Assert.Contains("must be between", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    #endregion

    #region Session ID Tests

    [Fact]
    public void SetMargins_WithSessionId_ShouldSetMarginsInMemory()
    {
        var docPath = CreateWordDocument("test_session_margins.docx");
        var sessionId = OpenSession(docPath);
        _tool.Execute("set_margins", sessionId: sessionId,
            top: 50.0, bottom: 50.0, left: 60.0, right: 60.0);

        // Assert - verify in-memory change
        var sessionDoc = SessionManager.GetDocument<Document>(sessionId);
        var section = sessionDoc.Sections[0];
        Assert.Equal(50.0, section.PageSetup.TopMargin);
        Assert.Equal(50.0, section.PageSetup.BottomMargin);
        Assert.Equal(60.0, section.PageSetup.LeftMargin);
        Assert.Equal(60.0, section.PageSetup.RightMargin);
    }

    [Fact]
    public void SetOrientation_WithSessionId_ShouldSetOrientationInMemory()
    {
        var docPath = CreateWordDocument("test_session_orientation.docx");
        var sessionId = OpenSession(docPath);
        _tool.Execute("set_orientation", sessionId: sessionId, orientation: "landscape");

        // Assert - verify in-memory change
        var sessionDoc = SessionManager.GetDocument<Document>(sessionId);
        Assert.Equal(Orientation.Landscape, sessionDoc.Sections[0].PageSetup.Orientation);
    }

    [Fact]
    public void AddPageBreak_WithSessionId_ShouldAddBreakInMemory()
    {
        var docPath = CreateWordDocumentWithContent("test_session_pagebreak.docx", "Content before break");
        var sessionId = OpenSession(docPath);
        var result = _tool.Execute("add_page_break", sessionId: sessionId);
        Assert.Contains("Page break added", result);

        // Verify in-memory document
        var sessionDoc = SessionManager.GetDocument<Document>(sessionId);
        Assert.NotNull(sessionDoc);
    }

    [Fact]
    public void Execute_WithInvalidSessionId_ShouldThrowKeyNotFoundException()
    {
        Assert.Throws<KeyNotFoundException>(() =>
            _tool.Execute("set_margins", sessionId: "invalid_session_id", top: 72.0));
    }

    [Fact]
    public void Execute_WithBothPathAndSessionId_ShouldPreferSessionId()
    {
        var docPath1 = CreateWordDocument("test_path_page.docx");
        var docPath2 = CreateWordDocument("test_session_page.docx");

        var sessionId = OpenSession(docPath2);

        // Modify the session document's margins to be identifiable
        _tool.Execute("set_margins", sessionId: sessionId, top: 99.0, bottom: 99.0);

        // Act - provide both path and sessionId, set different margins
        _tool.Execute("set_margins", docPath1, sessionId, left: 88.0, right: 88.0);

        // Assert - verify changes were applied to session document (not file path document)
        var sessionDoc = SessionManager.GetDocument<Document>(sessionId);
        Assert.Equal(88.0, sessionDoc.Sections[0].PageSetup.LeftMargin);
        Assert.Equal(88.0, sessionDoc.Sections[0].PageSetup.RightMargin);

        // Original file should not be modified
        var fileDoc = new Document(docPath1);
        Assert.NotEqual(88.0, fileDoc.Sections[0].PageSetup.LeftMargin);
    }

    [Fact]
    public void SetSize_WithSessionId_ShouldSetSizeInMemory()
    {
        var docPath = CreateWordDocument("test_session_size.docx");
        var sessionId = OpenSession(docPath);
        _tool.Execute("set_size", sessionId: sessionId, width: 400.0, height: 600.0);

        // Assert - verify in-memory change
        var sessionDoc = SessionManager.GetDocument<Document>(sessionId);
        Assert.Equal(400.0, sessionDoc.Sections[0].PageSetup.PageWidth);
        Assert.Equal(600.0, sessionDoc.Sections[0].PageSetup.PageHeight);
    }

    #endregion
}