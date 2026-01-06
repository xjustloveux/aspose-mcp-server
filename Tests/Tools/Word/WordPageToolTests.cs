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

    #region General

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
    public void SetOrientation_WithPortrait_ShouldSetPortrait()
    {
        var docPath = CreateWordDocument("test_portrait.docx");
        var doc = new Document(docPath);
        doc.Sections[0].PageSetup.Orientation = Orientation.Landscape;
        doc.Save(docPath);

        var outputPath = CreateTestFilePath("test_portrait_output.docx");
        _tool.Execute("set_orientation", docPath, outputPath: outputPath, orientation: "portrait");

        var resultDoc = new Document(outputPath);
        Assert.Equal(Orientation.Portrait, resultDoc.Sections[0].PageSetup.Orientation);
    }

    [Fact]
    public void SetOrientation_WithInvalidOrientation_ShouldDefaultToPortrait()
    {
        var docPath = CreateWordDocument("test_invalid_orientation.docx");
        var outputPath = CreateTestFilePath("test_invalid_orientation_output.docx");

        var result = _tool.Execute("set_orientation", docPath, outputPath: outputPath, orientation: "diagonal");
        Assert.True(File.Exists(outputPath));
        Assert.StartsWith("Page orientation set to", result);

        var doc = new Document(outputPath);
        Assert.Equal(Orientation.Portrait, doc.FirstSection.PageSetup.Orientation);
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

    [Theory]
    [InlineData("A4", PaperSize.A4)]
    [InlineData("Letter", PaperSize.Letter)]
    public void SetSize_WithPaperSize_ShouldSetPredefinedSize(string paperSizeName, PaperSize expectedPaperSize)
    {
        var docPath = CreateWordDocument($"test_paper_size_{paperSizeName}.docx");
        var outputPath = CreateTestFilePath($"test_paper_size_{paperSizeName}_output.docx");
        _tool.Execute("set_size", docPath, outputPath: outputPath, paperSize: paperSizeName);

        var doc = new Document(outputPath);
        Assert.Equal(expectedPaperSize, doc.Sections[0].PageSetup.PaperSize);
    }

    [Fact]
    public void SetPageNumber_ShouldSetPageNumber()
    {
        var docPath = CreateWordDocument("test_set_page_number.docx");
        var outputPath = CreateTestFilePath("test_set_page_number_output.docx");
        _tool.Execute("set_page_number", docPath, outputPath: outputPath, startingPageNumber: 5);
        var doc = new Document(outputPath);
        var section = doc.Sections[0];
        Assert.True(section.PageSetup.RestartPageNumbering || section.PageSetup.PageStartingNumber == 5,
            "Page starting number should be set");
    }

    [Theory]
    [InlineData("roman", NumberStyle.UppercaseRoman)]
    [InlineData("letter", NumberStyle.UppercaseLetter)]
    public void SetPageNumber_WithFormat_ShouldSetFormat(string format, NumberStyle expectedStyle)
    {
        var docPath = CreateWordDocument($"test_{format}_format.docx");
        var outputPath = CreateTestFilePath($"test_{format}_format_output.docx");
        _tool.Execute("set_page_number", docPath, outputPath: outputPath, pageNumberFormat: format);

        var doc = new Document(outputPath);
        Assert.Equal(expectedStyle, doc.Sections[0].PageSetup.PageNumberStyle);
    }

    [Fact]
    public void SetPageSetup_WithAllOptions_ShouldSetAllOptions()
    {
        var docPath = CreateWordDocument("test_page_setup_all.docx");
        var outputPath = CreateTestFilePath("test_page_setup_all_output.docx");
        var result = _tool.Execute("set_page_setup", docPath, outputPath: outputPath,
            top: 50.0, bottom: 50.0, left: 60.0, right: 60.0, orientation: "landscape");

        Assert.StartsWith("Page setup updated:", result);

        var doc = new Document(outputPath);
        var pageSetup = doc.Sections[0].PageSetup;
        Assert.Equal(50.0, pageSetup.TopMargin);
        Assert.Equal(50.0, pageSetup.BottomMargin);
        Assert.Equal(60.0, pageSetup.LeftMargin);
        Assert.Equal(60.0, pageSetup.RightMargin);
        Assert.Equal(Orientation.Landscape, pageSetup.Orientation);
    }

    [Fact]
    public void DeletePage_ShouldRemoveSpecifiedPage()
    {
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
    public void DeletePage_FirstPage_ShouldDeleteFirstPage()
    {
        var docPath = CreateTestFilePath("test_delete_first_page.docx");
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.Writeln("Page 1 - to be deleted");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("Page 2 - should remain");
        doc.Save(docPath);

        var outputPath = CreateTestFilePath("test_delete_first_page_output.docx");
        var result = _tool.Execute("delete_page", docPath, outputPath: outputPath, pageIndex: 0);

        Assert.Contains("deleted successfully", result);
        var resultDoc = new Document(outputPath);
        Assert.DoesNotContain("to be deleted", resultDoc.GetText());
    }

    [Fact]
    public void InsertBlankPage_ShouldInsertPageAtSpecifiedPosition()
    {
        var docPath = CreateWordDocumentWithContent("test_insert_blank.docx", "Existing content");
        var outputPath = CreateTestFilePath("test_insert_blank_output.docx");
        var result = _tool.Execute("insert_blank_page", docPath, outputPath: outputPath);
        Assert.StartsWith("Blank page inserted", result);
        Assert.True(File.Exists(outputPath));
    }

    [Fact]
    public void InsertBlankPage_AtBeginning_ShouldInsertAtStart()
    {
        var docPath = CreateWordDocumentWithContent("test_insert_beginning.docx", "Existing content");
        var outputPath = CreateTestFilePath("test_insert_beginning_output.docx");
        var result = _tool.Execute("insert_blank_page", docPath, outputPath: outputPath, insertAtPageIndex: 0);

        Assert.StartsWith("Blank page inserted", result);
        Assert.True(File.Exists(outputPath));
    }

    [Fact]
    public void AddPageBreak_ShouldAddPageBreakAtDocumentEnd()
    {
        var docPath = CreateWordDocumentWithContent("test_add_page_break.docx", "Content before break");
        var outputPath = CreateTestFilePath("test_add_page_break_output.docx");
        var result = _tool.Execute("add_page_break", docPath, outputPath: outputPath);
        Assert.StartsWith("Page break added", result);
        var doc = new Document(outputPath);
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
        Assert.StartsWith("Page break added", result);
    }

    [Theory]
    [InlineData("SET_MARGINS")]
    [InlineData("Set_Margins")]
    [InlineData("set_margins")]
    public void Execute_SetMarginsOperationIsCaseInsensitive(string operation)
    {
        var docPath = CreateWordDocument($"test_case_{operation.Replace("_", "")}.docx");
        var outputPath = CreateTestFilePath($"test_case_{operation.Replace("_", "")}_output.docx");
        var result = _tool.Execute(operation, docPath, outputPath: outputPath, top: 72.0);
        Assert.StartsWith("Page margins updated", result);
    }

    [Theory]
    [InlineData("SET_ORIENTATION")]
    [InlineData("Set_Orientation")]
    [InlineData("set_orientation")]
    public void Execute_SetOrientationOperationIsCaseInsensitive(string operation)
    {
        var docPath = CreateWordDocument($"test_orient_{operation.Replace("_", "")}.docx");
        var outputPath = CreateTestFilePath($"test_orient_{operation.Replace("_", "")}_output.docx");
        var result = _tool.Execute(operation, docPath, outputPath: outputPath, orientation: "landscape");
        Assert.StartsWith("Page orientation set to", result);
    }

    [Theory]
    [InlineData("SET_SIZE")]
    [InlineData("Set_Size")]
    [InlineData("set_size")]
    public void Execute_SetSizeOperationIsCaseInsensitive(string operation)
    {
        var docPath = CreateWordDocument($"test_size_{operation.Replace("_", "")}.docx");
        var outputPath = CreateTestFilePath($"test_size_{operation.Replace("_", "")}_output.docx");
        var result = _tool.Execute(operation, docPath, outputPath: outputPath, width: 500.0, height: 700.0);
        Assert.StartsWith("Page size updated", result);
    }

    [Theory]
    [InlineData("SET_PAGE_NUMBER")]
    [InlineData("Set_Page_Number")]
    [InlineData("set_page_number")]
    public void Execute_SetPageNumberOperationIsCaseInsensitive(string operation)
    {
        var docPath = CreateWordDocument($"test_pagenum_{operation.Replace("_", "")}.docx");
        var outputPath = CreateTestFilePath($"test_pagenum_{operation.Replace("_", "")}_output.docx");
        var result = _tool.Execute(operation, docPath, outputPath: outputPath, startingPageNumber: 1);
        Assert.StartsWith("Page number settings updated", result);
    }

    [Theory]
    [InlineData("SET_PAGE_SETUP")]
    [InlineData("Set_Page_Setup")]
    [InlineData("set_page_setup")]
    public void Execute_SetPageSetupOperationIsCaseInsensitive(string operation)
    {
        var docPath = CreateWordDocument($"test_setup_{operation.Replace("_", "")}.docx");
        var outputPath = CreateTestFilePath($"test_setup_{operation.Replace("_", "")}_output.docx");
        var result = _tool.Execute(operation, docPath, outputPath: outputPath, top: 50.0);
        Assert.StartsWith("Page setup updated:", result);
    }

    [Theory]
    [InlineData("ADD_PAGE_BREAK")]
    [InlineData("Add_Page_Break")]
    [InlineData("add_page_break")]
    public void Execute_AddPageBreakOperationIsCaseInsensitive(string operation)
    {
        var docPath = CreateWordDocumentWithContent($"test_break_{operation.Replace("_", "")}.docx", "Content");
        var outputPath = CreateTestFilePath($"test_break_{operation.Replace("_", "")}_output.docx");
        var result = _tool.Execute(operation, docPath, outputPath: outputPath);
        Assert.StartsWith("Page break added", result);
    }

    [Theory]
    [InlineData("INSERT_BLANK_PAGE")]
    [InlineData("Insert_Blank_Page")]
    [InlineData("insert_blank_page")]
    public void Execute_InsertBlankPageOperationIsCaseInsensitive(string operation)
    {
        var docPath = CreateWordDocumentWithContent($"test_insert_{operation.Replace("_", "")}.docx", "Content");
        var outputPath = CreateTestFilePath($"test_insert_{operation.Replace("_", "")}_output.docx");
        var result = _tool.Execute(operation, docPath, outputPath: outputPath);
        Assert.StartsWith("Blank page inserted", result);
    }

    [Theory]
    [InlineData("DELETE_PAGE")]
    [InlineData("Delete_Page")]
    [InlineData("delete_page")]
    public void Execute_DeletePageOperationIsCaseInsensitive(string operation)
    {
        var docPath = CreateTestFilePath($"test_delete_{operation.Replace("_", "")}.docx");
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.Writeln("Page 1");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("Page 2");
        doc.Save(docPath);

        var outputPath = CreateTestFilePath($"test_delete_{operation.Replace("_", "")}_output.docx");
        var result = _tool.Execute(operation, docPath, outputPath: outputPath, pageIndex: 1);
        Assert.Contains("deleted successfully", result);
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
    public void DeletePage_WithInvalidPageIndex_ShouldThrowArgumentException()
    {
        var docPath = CreateWordDocumentWithContent("test_delete_invalid.docx", "Single page content");
        var outputPath = CreateTestFilePath("test_delete_invalid_output.docx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("delete_page", docPath, outputPath: outputPath, pageIndex: 999));

        Assert.Contains("must be between", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void DeletePage_WithoutPageIndex_ShouldThrowArgumentException()
    {
        var docPath = CreateWordDocumentWithContent("test_delete_no_index.docx", "Content");
        var outputPath = CreateTestFilePath("test_delete_no_index_output.docx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("delete_page", docPath, outputPath: outputPath));

        Assert.Contains("pageIndex parameter is required", ex.Message);
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

    [Fact]
    public void SetOrientation_WithoutOrientation_ShouldThrowArgumentException()
    {
        var docPath = CreateWordDocument("test_orientation_missing.docx");
        var outputPath = CreateTestFilePath("test_orientation_missing_output.docx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("set_orientation", docPath, outputPath: outputPath));

        Assert.Contains("orientation parameter is required", ex.Message);
    }

    [Fact]
    public void SetSize_WithoutPaperSizeOrDimensions_ShouldThrowArgumentException()
    {
        var docPath = CreateWordDocument("test_size_missing.docx");
        var outputPath = CreateTestFilePath("test_size_missing_output.docx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("set_size", docPath, outputPath: outputPath));

        Assert.Contains("Either paperSize or both width and height must be provided", ex.Message);
    }

    [Fact]
    public void SetSize_WithOnlyWidth_ShouldThrowArgumentException()
    {
        var docPath = CreateWordDocument("test_size_only_width.docx");
        var outputPath = CreateTestFilePath("test_size_only_width_output.docx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("set_size", docPath, outputPath: outputPath, width: 500.0));

        Assert.Contains("Either paperSize or both width and height must be provided", ex.Message);
    }

    [Fact]
    public void SetMargins_WithInvalidSectionIndex_ShouldThrowArgumentException()
    {
        var docPath = CreateWordDocument("test_margins_invalid_section.docx");
        var outputPath = CreateTestFilePath("test_margins_invalid_section_output.docx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("set_margins", docPath, outputPath: outputPath, top: 72.0, sectionIndex: 999));

        Assert.Contains("sectionIndex", ex.Message);
    }

    [Fact]
    public void SetPageSetup_WithInvalidSectionIndex_ShouldThrowArgumentException()
    {
        var docPath = CreateWordDocument("test_page_setup_invalid_section.docx");
        var outputPath = CreateTestFilePath("test_page_setup_invalid_section_output.docx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("set_page_setup", docPath, outputPath: outputPath, top: 72.0, sectionIndex: 999));

        Assert.Contains("sectionIndex must be between", ex.Message);
    }

    [Fact]
    public void InsertBlankPage_WithInvalidPageIndex_ShouldThrowArgumentException()
    {
        var docPath = CreateWordDocumentWithContent("test_insert_invalid_page.docx", "Content");
        var outputPath = CreateTestFilePath("test_insert_invalid_page_output.docx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("insert_blank_page", docPath, outputPath: outputPath, insertAtPageIndex: 999));

        Assert.Contains("insertAtPageIndex must be between", ex.Message);
    }

    #endregion

    #region Session

    [Fact]
    public void SetMargins_WithSessionId_ShouldSetMarginsInMemory()
    {
        var docPath = CreateWordDocument("test_session_margins.docx");
        var sessionId = OpenSession(docPath);
        _tool.Execute("set_margins", sessionId: sessionId,
            top: 50.0, bottom: 50.0, left: 60.0, right: 60.0);

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

        var sessionDoc = SessionManager.GetDocument<Document>(sessionId);
        Assert.Equal(Orientation.Landscape, sessionDoc.Sections[0].PageSetup.Orientation);
    }

    [Fact]
    public void SetSize_WithSessionId_ShouldSetSizeInMemory()
    {
        var docPath = CreateWordDocument("test_session_size.docx");
        var sessionId = OpenSession(docPath);
        _tool.Execute("set_size", sessionId: sessionId, width: 400.0, height: 600.0);

        var sessionDoc = SessionManager.GetDocument<Document>(sessionId);
        Assert.Equal(400.0, sessionDoc.Sections[0].PageSetup.PageWidth);
        Assert.Equal(600.0, sessionDoc.Sections[0].PageSetup.PageHeight);
    }

    [Fact]
    public void SetPageNumber_WithSessionId_ShouldSetPageNumberInMemory()
    {
        var docPath = CreateWordDocument("test_session_page_number.docx");
        var sessionId = OpenSession(docPath);
        _tool.Execute("set_page_number", sessionId: sessionId, startingPageNumber: 10);

        var sessionDoc = SessionManager.GetDocument<Document>(sessionId);
        Assert.Equal(10, sessionDoc.Sections[0].PageSetup.PageStartingNumber);
        Assert.True(sessionDoc.Sections[0].PageSetup.RestartPageNumbering);
    }

    [Fact]
    public void SetPageSetup_WithSessionId_ShouldSetPageSetupInMemory()
    {
        var docPath = CreateWordDocument("test_session_page_setup.docx");
        var sessionId = OpenSession(docPath);
        var result = _tool.Execute("set_page_setup", sessionId: sessionId,
            top: 36.0, bottom: 36.0, orientation: "landscape");

        Assert.StartsWith("Page setup updated:", result);
        var sessionDoc = SessionManager.GetDocument<Document>(sessionId);
        Assert.Equal(36.0, sessionDoc.Sections[0].PageSetup.TopMargin);
        Assert.Equal(Orientation.Landscape, sessionDoc.Sections[0].PageSetup.Orientation);
    }

    [Fact]
    public void AddPageBreak_WithSessionId_ShouldAddBreakInMemory()
    {
        var docPath = CreateWordDocumentWithContent("test_session_pagebreak.docx", "Content before break");
        var sessionId = OpenSession(docPath);
        var result = _tool.Execute("add_page_break", sessionId: sessionId);
        Assert.StartsWith("Page break added", result);

        var sessionDoc = SessionManager.GetDocument<Document>(sessionId);
        Assert.NotNull(sessionDoc);
    }

    [Fact]
    public void DeletePage_WithSessionId_ShouldDeletePageInMemory()
    {
        var docPath = CreateTestFilePath("test_session_delete_page.docx");
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.Writeln("Page 1 content");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("Page 2 content");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("Page 3 content");
        doc.Save(docPath);

        var sessionId = OpenSession(docPath);
        var result = _tool.Execute("delete_page", sessionId: sessionId, pageIndex: 1);

        Assert.Contains("deleted successfully", result);
    }

    [Fact]
    public void InsertBlankPage_WithSessionId_ShouldInsertPageInMemory()
    {
        var docPath = CreateWordDocumentWithContent("test_session_insert_blank.docx", "Content");
        var sessionId = OpenSession(docPath);
        var result = _tool.Execute("insert_blank_page", sessionId: sessionId);

        Assert.StartsWith("Blank page inserted", result);
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
        Assert.Equal(88.0, sessionDoc.Sections[0].PageSetup.RightMargin);

        var fileDoc = new Document(docPath1);
        Assert.NotEqual(88.0, fileDoc.Sections[0].PageSetup.LeftMargin);
    }

    #endregion
}