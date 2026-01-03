using Aspose.Words;
using AsposeMcpServer.Tests.Helpers;
using AsposeMcpServer.Tools.Word;

namespace AsposeMcpServer.Tests.Tools.Word;

public class WordSectionToolTests : WordTestBase
{
    private readonly WordSectionTool _tool;

    public WordSectionToolTests()
    {
        _tool = new WordSectionTool(SessionManager);
    }

    #region General Tests

    [Fact]
    public void InsertSection_ShouldInsertSection()
    {
        var docPath = CreateWordDocumentWithContent("test_insert_section.docx", "Content before section");
        var outputPath = CreateTestFilePath("test_insert_section_output.docx");
        _tool.Execute("insert", docPath, outputPath: outputPath,
            sectionBreakType: "NextPage", insertAtParagraphIndex: 0);
        var doc = new Document(outputPath);
        Assert.True(doc.Sections.Count > 1, "Document should contain multiple sections");
    }

    [Fact]
    public void InsertSection_AtDocumentEnd_ShouldWork()
    {
        var docPath = CreateWordDocumentWithContent("test_insert_end.docx", "Content");
        var outputPath = CreateTestFilePath("test_insert_end_output.docx");
        var result = _tool.Execute("insert", docPath, outputPath: outputPath,
            sectionBreakType: "Continuous", insertAtParagraphIndex: -1);
        Assert.Contains("Section break inserted", result);
        Assert.Contains("Continuous", result);
    }

    [Fact]
    public void InsertSection_WithInvalidSectionIndex_ShouldThrowException()
    {
        var docPath = CreateWordDocumentWithContent("test_insert_invalid_sec.docx", "Content");
        var exception = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("insert", docPath,
                sectionBreakType: "NextPage", sectionIndex: 99, insertAtParagraphIndex: 0));
        Assert.Contains("sectionIndex must be between", exception.Message);
    }

    [Fact]
    public void InsertSection_WithInvalidParagraphIndex_ShouldThrowException()
    {
        var docPath = CreateWordDocumentWithContent("test_insert_invalid_para.docx", "Content");
        var exception = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("insert", docPath,
                sectionBreakType: "NextPage", insertAtParagraphIndex: 999));
        Assert.Contains("insertAtParagraphIndex must be between", exception.Message);
    }

    [Fact]
    public void GetSections_ShouldReturnSectionsInfo()
    {
        var docPath = CreateWordDocument("test_get_sections.docx");
        var doc = new Document(docPath);
        var builder = new DocumentBuilder(doc);
        builder.InsertBreak(BreakType.PageBreak);
        builder.CurrentSection.PageSetup.SectionStart = SectionStart.NewPage;
        doc.Save(docPath);
        var result = _tool.Execute("get", docPath);
        Assert.NotNull(result);
        Assert.NotEmpty(result);
        Assert.Contains("\"sections\"", result); // JSON format
        Assert.Contains("\"sectionBreak\"", result);
        Assert.Contains("\"type\"", result);
    }

    [Fact]
    public void GetSections_WithSpecificIndex_ShouldReturnSingleSection()
    {
        var docPath = CreateWordDocument("test_get_single_section.docx");
        var doc = new Document(docPath);
        var builder = new DocumentBuilder(doc);
        builder.InsertBreak(BreakType.SectionBreakNewPage);
        builder.Writeln("Second section content");
        doc.Save(docPath);
        var result = _tool.Execute("get", docPath, sectionIndex: 0);
        Assert.Contains("\"section\"", result); // Single section object
        Assert.Contains("\"index\": 0", result);
        Assert.DoesNotContain("\"index\": 1", result);
    }

    [Fact]
    public void GetSections_WithInvalidIndex_ShouldThrowException()
    {
        var docPath = CreateWordDocumentWithContent("test_get_invalid_index.docx", "Content");
        var exception = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("get", docPath, sectionIndex: 99));
        Assert.Contains("sectionIndex must be between", exception.Message);
    }

    [Fact]
    public void DeleteSection_ShouldDeleteSection()
    {
        var docPath = CreateWordDocument("test_delete_section.docx");
        var doc = new Document(docPath);
        var builder = new DocumentBuilder(doc);
        builder.InsertBreak(BreakType.SectionBreakNewPage);
        doc.Save(docPath);

        doc = new Document(docPath);
        var sectionsBefore = doc.Sections.Count;
        Assert.True(sectionsBefore > 1, "Document should have multiple sections before deletion");

        var outputPath = CreateTestFilePath("test_delete_section_output.docx");
        var result = _tool.Execute("delete", docPath, outputPath: outputPath, sectionIndex: 1);
        var resultDoc = new Document(outputPath);
        Assert.Equal(sectionsBefore - 1, resultDoc.Sections.Count);
        Assert.Contains("Deleted", result);
        Assert.Contains("with their content", result);
    }

    [Fact]
    public void DeleteSection_LastSection_ShouldThrowException()
    {
        var docPath = CreateWordDocumentWithContent("test_delete_last.docx", "Single section");
        var exception = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("delete", docPath, sectionIndex: 0));
        Assert.Contains("Cannot delete the last section", exception.Message);
    }

    [Fact]
    public void DeleteSection_MultipleSections_ShouldDeleteAll()
    {
        var docPath = CreateWordDocument("test_delete_multiple.docx");
        var doc = new Document(docPath);
        var builder = new DocumentBuilder(doc);
        builder.InsertBreak(BreakType.SectionBreakNewPage);
        builder.Writeln("Section 2");
        builder.InsertBreak(BreakType.SectionBreakNewPage);
        builder.Writeln("Section 3");
        doc.Save(docPath);

        var outputPath = CreateTestFilePath("test_delete_multiple_output.docx");
        var result = _tool.Execute("delete", docPath, outputPath: outputPath, sectionIndices: [1, 2]);
        var resultDoc = new Document(outputPath);
        Assert.Equal(1, resultDoc.Sections.Count);
        Assert.Contains("Deleted 2 section(s)", result);
    }

    [Fact]
    public void DeleteSection_WithoutIndex_ShouldThrowException()
    {
        var docPath = CreateWordDocument("test_delete_no_index.docx");
        var doc = new Document(docPath);
        var builder = new DocumentBuilder(doc);
        builder.InsertBreak(BreakType.SectionBreakNewPage);
        doc.Save(docPath);
        var exception = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("delete", docPath));
        Assert.Contains("sectionIndex or sectionIndices must be provided", exception.Message);
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
    public void InsertSection_WithMissingSectionBreakType_ShouldThrowArgumentException()
    {
        var docPath = CreateWordDocumentWithContent("test_missing_break_type.docx", "Test content");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("insert", docPath, sectionBreakType: "", insertAtParagraphIndex: 0));

        Assert.Contains("sectionBreakType", ex.Message);
    }

    [Fact]
    public void InsertSection_WithInvalidSectionBreakType_ShouldDefaultToNextPage()
    {
        var docPath = CreateWordDocumentWithContent("test_invalid_break_type.docx", "Test content");
        var outputPath = CreateTestFilePath("test_invalid_break_type_output.docx");

        // Act - Invalid sectionBreakType defaults to NewPage (NextPage)
        var result = _tool.Execute("insert", docPath, outputPath: outputPath,
            sectionBreakType: "InvalidType", insertAtParagraphIndex: 0);
        Assert.Contains("Section break inserted", result);
        Assert.Contains("InvalidType", result);
        Assert.True(File.Exists(outputPath));

        // Verify the section was added
        var doc = new Document(outputPath);
        Assert.True(doc.Sections.Count >= 2);
    }

    #endregion

    #region Session ID Tests

    [Fact]
    public void GetSections_WithSessionId_ShouldReturnSections()
    {
        var docPath = CreateWordDocument("test_session_get_sections.docx");
        var doc = new Document(docPath);
        var builder = new DocumentBuilder(doc);
        builder.InsertBreak(BreakType.SectionBreakNewPage);
        builder.Writeln("Section 2");
        doc.Save(docPath);

        var sessionId = OpenSession(docPath);
        var result = _tool.Execute("get", sessionId: sessionId);
        Assert.Contains("\"sections\"", result);
    }

    [Fact]
    public void InsertSection_WithSessionId_ShouldInsertInMemory()
    {
        var docPath = CreateWordDocumentWithContent("test_session_insert.docx", "Content");
        var sessionId = OpenSession(docPath);
        var result = _tool.Execute("insert", sessionId: sessionId,
            sectionBreakType: "NextPage", insertAtParagraphIndex: 0);
        Assert.Contains("Section break inserted", result);

        // Verify in-memory document has the new section
        var sessionDoc = SessionManager.GetDocument<Document>(sessionId);
        Assert.True(sessionDoc.Sections.Count > 1);
    }

    [Fact]
    public void DeleteSection_WithSessionId_ShouldDeleteInMemory()
    {
        var docPath = CreateWordDocument("test_session_delete.docx");
        var doc = new Document(docPath);
        var builder = new DocumentBuilder(doc);
        builder.InsertBreak(BreakType.SectionBreakNewPage);
        builder.Writeln("Section 2");
        doc.Save(docPath);

        var sessionId = OpenSession(docPath);
        var result = _tool.Execute("delete", sessionId: sessionId, sectionIndex: 1);
        Assert.Contains("Deleted", result);

        // Verify in-memory document has one less section
        var sessionDoc = SessionManager.GetDocument<Document>(sessionId);
        Assert.Equal(1, sessionDoc.Sections.Count);
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
        var docPath1 = CreateWordDocumentWithContent("test_path_section.docx", "Path document");
        var docPath2 = CreateWordDocument("test_session_section.docx");
        var doc2 = new Document(docPath2);
        var builder = new DocumentBuilder(doc2);
        builder.InsertBreak(BreakType.SectionBreakNewPage);
        builder.Writeln("Session section 2");
        doc2.Save(docPath2);

        var sessionId = OpenSession(docPath2);

        // Act - provide both path and sessionId
        var result = _tool.Execute("get", docPath1, sessionId);

        // Assert - should use sessionId (which has 2 sections)
        Assert.Contains("\"sections\"", result);
        Assert.Contains("\"index\": 1", result); // Should have section index 1
    }

    #endregion
}