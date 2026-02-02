using Aspose.Words;
using AsposeMcpServer.Results;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;
using AsposeMcpServer.Tools.Word;

namespace AsposeMcpServer.Tests.Tools.Word;

/// <summary>
///     Integration tests for WordPageTool.
///     Focuses on session management, file I/O, and operation routing.
///     Detailed parameter validation and business logic tests are in Handler tests.
/// </summary>
public class WordPageToolTests : WordTestBase
{
    private readonly WordPageTool _tool;

    public WordPageToolTests()
    {
        _tool = new WordPageTool(SessionManager);
    }

    #region File I/O Smoke Tests

    [Fact]
    public void SetMargins_ShouldSetPageMarginsAndPersistToFile()
    {
        var docPath = CreateWordDocument("test_set_margins.docx");
        var outputPath = CreateTestFilePath("test_set_margins_output.docx");
        _tool.Execute("set_margins", docPath, outputPath: outputPath,
            top: 72.0, bottom: 72.0, left: 90.0, right: 90.0);
        var doc = new Document(outputPath);
        var section = doc.Sections[0];
        Assert.Equal(72.0, section.PageSetup.TopMargin);
        Assert.Equal(72.0, section.PageSetup.BottomMargin);
    }

    [Fact]
    public void SetOrientation_ShouldSetOrientationAndPersistToFile()
    {
        var docPath = CreateWordDocument("test_set_orientation.docx");
        var outputPath = CreateTestFilePath("test_set_orientation_output.docx");
        _tool.Execute("set_orientation", docPath, outputPath: outputPath, orientation: "landscape");
        var doc = new Document(outputPath);
        Assert.Equal(Orientation.Landscape, doc.Sections[0].PageSetup.Orientation);
    }

    [Fact]
    public void SetPageSize_ShouldSetSizeAndPersistToFile()
    {
        var docPath = CreateWordDocument("test_set_size.docx");
        var outputPath = CreateTestFilePath("test_set_size_output.docx");
        _tool.Execute("set_size", docPath, outputPath: outputPath, width: 595.0, height: 842.0);
        var doc = new Document(outputPath);
        Assert.Equal(595.0, doc.Sections[0].PageSetup.PageWidth);
        Assert.Equal(842.0, doc.Sections[0].PageSetup.PageHeight);
    }

    [Fact]
    public void SetPageNumber_ShouldSetPageNumberAndPersistToFile()
    {
        var docPath = CreateWordDocument("test_set_page_number.docx");
        var outputPath = CreateTestFilePath("test_set_page_number_output.docx");
        _tool.Execute("set_page_number", docPath, outputPath: outputPath, startingPageNumber: 5);
        var doc = new Document(outputPath);
        Assert.True(doc.Sections[0].PageSetup.RestartPageNumbering ||
                    doc.Sections[0].PageSetup.PageStartingNumber == 5);
    }

    [Fact]
    public void SetPageSetup_ShouldSetAllOptionsAndPersistToFile()
    {
        var docPath = CreateWordDocument("test_page_setup.docx");
        var outputPath = CreateTestFilePath("test_page_setup_output.docx");
        var result = _tool.Execute("set_page_setup", docPath, outputPath: outputPath,
            top: 50.0, bottom: 50.0, orientation: "landscape");
        Assert.IsType<FinalizedResult<SuccessResult>>(result);
        var doc = new Document(outputPath);
        Assert.Equal(50.0, doc.Sections[0].PageSetup.TopMargin);
        Assert.Equal(50.0, doc.Sections[0].PageSetup.BottomMargin);
        Assert.Equal(Orientation.Landscape, doc.Sections[0].PageSetup.Orientation);
    }

    [Fact]
    public void DeletePage_ShouldRemovePageAndPersistToFile()
    {
        var docPath = CreateTestFilePath("test_delete_page.docx");
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.Writeln("Page 1");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("Page 2");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("Page 3");
        doc.Save(docPath);

        var pageCountBefore = doc.PageCount;
        var outputPath = CreateTestFilePath("test_delete_page_output.docx");
        _tool.Execute("delete_page", docPath, outputPath: outputPath, pageIndex: 1);
        var resultDoc = new Document(outputPath);
        Assert.True(resultDoc.PageCount < pageCountBefore);
    }

    [Fact]
    public void InsertBlankPage_ShouldInsertPageAndPersistToFile()
    {
        var docPath = CreateWordDocumentWithContent("test_insert_blank.docx", "Existing content");
        var docBefore = new Document(docPath);
        var pageCountBefore = docBefore.PageCount;
        var outputPath = CreateTestFilePath("test_insert_blank_output.docx");
        _tool.Execute("insert_blank_page", docPath, outputPath: outputPath);
        Assert.True(File.Exists(outputPath));
        var resultDoc = new Document(outputPath);
        Assert.True(resultDoc.PageCount > pageCountBefore);
    }

    [Fact]
    public void AddPageBreak_ShouldAddBreakAndPersistToFile()
    {
        var docPath = CreateWordDocumentWithContent("test_add_page_break.docx", "Content before break");
        var outputPath = CreateTestFilePath("test_add_page_break_output.docx");
        var result = _tool.Execute("add_page_break", docPath, outputPath: outputPath);
        Assert.IsType<FinalizedResult<SuccessResult>>(result);
        Assert.True(File.Exists(outputPath));
    }

    #endregion

    #region Operation Routing

    [Theory]
    [InlineData("SET_MARGINS")]
    [InlineData("Set_Margins")]
    [InlineData("set_margins")]
    public void Operation_ShouldBeCaseInsensitive(string operation)
    {
        var docPath = CreateWordDocument($"test_case_{operation.Replace("_", "")}.docx");
        var outputPath = CreateTestFilePath($"test_case_{operation.Replace("_", "")}_output.docx");
        _tool.Execute(operation, docPath, outputPath: outputPath, top: 72.0);
        var doc = new Document(outputPath);
        Assert.Equal(72.0, doc.Sections[0].PageSetup.TopMargin);
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
        Assert.ThrowsAny<Exception>(() => _tool.Execute("set_margins", top: 72.0));
    }

    #endregion

    #region Session Management

    [Fact]
    public void SetMargins_WithSessionId_ShouldSetMarginsInMemory()
    {
        var docPath = CreateWordDocument("test_session_margins.docx");
        var sessionId = OpenSession(docPath);
        var result = _tool.Execute("set_margins", sessionId: sessionId,
            top: 50.0, bottom: 50.0, left: 60.0, right: 60.0);

        var output = GetResultOutput<SuccessResult>(result);
        Assert.True(output.IsSession);
        var sessionDoc = SessionManager.GetDocument<Document>(sessionId);
        Assert.Equal(50.0, sessionDoc.Sections[0].PageSetup.TopMargin);
        Assert.Equal(60.0, sessionDoc.Sections[0].PageSetup.LeftMargin);
    }

    [Fact]
    public void SetOrientation_WithSessionId_ShouldSetOrientationInMemory()
    {
        var docPath = CreateWordDocument("test_session_orientation.docx");
        var sessionId = OpenSession(docPath);
        var result = _tool.Execute("set_orientation", sessionId: sessionId, orientation: "landscape");

        var output = GetResultOutput<SuccessResult>(result);
        Assert.True(output.IsSession);
        var sessionDoc = SessionManager.GetDocument<Document>(sessionId);
        Assert.Equal(Orientation.Landscape, sessionDoc.Sections[0].PageSetup.Orientation);
    }

    [Fact]
    public void SetSize_WithSessionId_ShouldSetSizeInMemory()
    {
        var docPath = CreateWordDocument("test_session_size.docx");
        var sessionId = OpenSession(docPath);
        var result = _tool.Execute("set_size", sessionId: sessionId, width: 400.0, height: 600.0);

        var output = GetResultOutput<SuccessResult>(result);
        Assert.True(output.IsSession);
        var sessionDoc = SessionManager.GetDocument<Document>(sessionId);
        Assert.Equal(400.0, sessionDoc.Sections[0].PageSetup.PageWidth);
        Assert.Equal(600.0, sessionDoc.Sections[0].PageSetup.PageHeight);
    }

    [Fact]
    public void SetPageNumber_WithSessionId_ShouldSetPageNumberInMemory()
    {
        var docPath = CreateWordDocument("test_session_page_number.docx");
        var sessionId = OpenSession(docPath);
        var result = _tool.Execute("set_page_number", sessionId: sessionId, startingPageNumber: 10);

        var output = GetResultOutput<SuccessResult>(result);
        Assert.True(output.IsSession);
        var sessionDoc = SessionManager.GetDocument<Document>(sessionId);
        Assert.Equal(10, sessionDoc.Sections[0].PageSetup.PageStartingNumber);
    }

    [Fact]
    public void SetPageSetup_WithSessionId_ShouldSetSetupInMemory()
    {
        var docPath = CreateWordDocument("test_session_page_setup.docx");
        var sessionId = OpenSession(docPath);
        var result = _tool.Execute("set_page_setup", sessionId: sessionId,
            top: 36.0, bottom: 36.0, orientation: "landscape");

        var output = GetResultOutput<SuccessResult>(result);
        Assert.True(output.IsSession);
        var sessionDoc = SessionManager.GetDocument<Document>(sessionId);
        Assert.Equal(36.0, sessionDoc.Sections[0].PageSetup.TopMargin);
        Assert.Equal(36.0, sessionDoc.Sections[0].PageSetup.BottomMargin);
        Assert.Equal(Orientation.Landscape, sessionDoc.Sections[0].PageSetup.Orientation);
    }

    [Fact]
    public void AddPageBreak_WithSessionId_ShouldAddBreakInMemory()
    {
        var docPath = CreateWordDocumentWithContent("test_session_pagebreak.docx", "Content before break");
        var sessionId = OpenSession(docPath);
        var result = _tool.Execute("add_page_break", sessionId: sessionId);
        var output = GetResultOutput<SuccessResult>(result);
        Assert.True(output.IsSession);
    }

    [Fact]
    public void DeletePage_WithSessionId_ShouldDeletePageInMemory()
    {
        var docPath = CreateTestFilePath("test_session_delete_page.docx");
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.Writeln("Page 1");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("Page 2");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("Page 3");
        doc.Save(docPath);

        var sessionId = OpenSession(docPath);
        var sessionDoc = SessionManager.GetDocument<Document>(sessionId);
        sessionDoc.UpdatePageLayout();
        var pageCountBefore = sessionDoc.PageCount;
        var result = _tool.Execute("delete_page", sessionId: sessionId, pageIndex: 1);
        var output = GetResultOutput<SuccessResult>(result);
        Assert.True(output.IsSession);
        sessionDoc = SessionManager.GetDocument<Document>(sessionId);
        sessionDoc.UpdatePageLayout();
        Assert.True(sessionDoc.PageCount < pageCountBefore);
    }

    [Fact]
    public void InsertBlankPage_WithSessionId_ShouldInsertPageInMemory()
    {
        var docPath = CreateWordDocumentWithContent("test_session_insert_blank.docx", "Content");
        var sessionId = OpenSession(docPath);
        var sessionDoc = SessionManager.GetDocument<Document>(sessionId);
        sessionDoc.UpdatePageLayout();
        var pageCountBefore = sessionDoc.PageCount;
        var result = _tool.Execute("insert_blank_page", sessionId: sessionId);
        var output = GetResultOutput<SuccessResult>(result);
        Assert.True(output.IsSession);
        sessionDoc = SessionManager.GetDocument<Document>(sessionId);
        sessionDoc.UpdatePageLayout();
        Assert.True(sessionDoc.PageCount > pageCountBefore);
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
        _tool.Execute("set_margins", sessionId: sessionId, top: 99.0, bottom: 99.0);
        _tool.Execute("set_margins", docPath1, sessionId, left: 88.0, right: 88.0);

        var sessionDoc = SessionManager.GetDocument<Document>(sessionId);
        Assert.Equal(88.0, sessionDoc.Sections[0].PageSetup.LeftMargin);

        var fileDoc = new Document(docPath1);
        Assert.NotEqual(88.0, fileDoc.Sections[0].PageSetup.LeftMargin);
    }

    #endregion
}
