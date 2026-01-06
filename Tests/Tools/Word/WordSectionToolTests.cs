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

    #region General

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
        Assert.StartsWith("Section break inserted", result);
        Assert.Contains("Continuous", result);
    }

    [Fact]
    public void InsertSection_WithNegativeParagraphIndex_ShouldInsertAtEnd()
    {
        var docPath = CreateWordDocumentWithContent("test_neg_para_idx.docx", "Content");
        var outputPath = CreateTestFilePath("test_neg_para_idx_output.docx");
        var result = _tool.Execute("insert", docPath, outputPath: outputPath,
            sectionBreakType: "NextPage", insertAtParagraphIndex: -1);
        Assert.StartsWith("Section break inserted", result);
    }

    [Theory]
    [InlineData("NextPage")]
    [InlineData("Continuous")]
    [InlineData("EvenPage")]
    [InlineData("OddPage")]
    public void InsertSection_WithDifferentBreakTypes_ShouldWork(string breakType)
    {
        var docPath = CreateWordDocumentWithContent($"test_{breakType.ToLower()}.docx", "Content");
        var outputPath = CreateTestFilePath($"test_{breakType.ToLower()}_output.docx");
        var result = _tool.Execute("insert", docPath, outputPath: outputPath,
            sectionBreakType: breakType, insertAtParagraphIndex: 0);
        Assert.Contains(breakType, result);
        var doc = new Document(outputPath);
        Assert.True(doc.Sections.Count > 1);
    }

    [Fact]
    public void InsertSection_WithInvalidSectionBreakType_ShouldDefaultToNextPage()
    {
        var docPath = CreateWordDocumentWithContent("test_invalid_break_type.docx", "Test content");
        var outputPath = CreateTestFilePath("test_invalid_break_type_output.docx");

        var result = _tool.Execute("insert", docPath, outputPath: outputPath,
            sectionBreakType: "InvalidType", insertAtParagraphIndex: 0);
        Assert.StartsWith("Section break inserted", result);
        Assert.Contains("InvalidType", result);
        Assert.True(File.Exists(outputPath));

        var doc = new Document(outputPath);
        Assert.True(doc.Sections.Count >= 2);
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
        Assert.Contains("\"sections\"", result);
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
        Assert.Contains("\"section\"", result);
        Assert.Contains("\"index\": 0", result);
        Assert.DoesNotContain("\"index\": 1", result);
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
        Assert.StartsWith("Deleted", result);
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
        Assert.StartsWith("Deleted", result);
    }

    [Fact]
    public void DeleteSection_WithNegativeSectionIndex_ShouldNotDelete()
    {
        var docPath = CreateWordDocument("test_delete_neg_idx.docx");
        var doc = new Document(docPath);
        var builder = new DocumentBuilder(doc);
        builder.InsertBreak(BreakType.SectionBreakNewPage);
        builder.Writeln("Section 2");
        doc.Save(docPath);

        var outputPath = CreateTestFilePath("test_delete_neg_idx_output.docx");
        var result = _tool.Execute("delete", docPath, outputPath: outputPath, sectionIndex: -1);
        Assert.Contains("Deleted 0 section(s)", result);
    }

    [Fact]
    public void DeleteSection_WithMixedValidInvalidIndices_ShouldDeleteOnlyValid()
    {
        var docPath = CreateWordDocument("test_delete_mixed_idx.docx");
        var doc = new Document(docPath);
        var builder = new DocumentBuilder(doc);
        builder.InsertBreak(BreakType.SectionBreakNewPage);
        builder.Writeln("Section 2");
        builder.InsertBreak(BreakType.SectionBreakNewPage);
        builder.Writeln("Section 3");
        doc.Save(docPath);

        var outputPath = CreateTestFilePath("test_delete_mixed_idx_output.docx");
        var result = _tool.Execute("delete", docPath, outputPath: outputPath,
            sectionIndices: [1, 99, -1]);
        Assert.StartsWith("Deleted", result);
    }

    [Fact]
    public void DeleteSection_AllButOne_ShouldStopAtOne()
    {
        var docPath = CreateWordDocument("test_delete_all.docx");
        var doc = new Document(docPath);
        var builder = new DocumentBuilder(doc);
        builder.InsertBreak(BreakType.SectionBreakNewPage);
        builder.Writeln("Section 2");
        builder.InsertBreak(BreakType.SectionBreakNewPage);
        builder.Writeln("Section 3");
        doc.Save(docPath);

        var outputPath = CreateTestFilePath("test_delete_all_output.docx");
        _ = _tool.Execute("delete", docPath, outputPath: outputPath,
            sectionIndices: [0, 1, 2]);
        var resultDoc = new Document(outputPath);
        Assert.Equal(1, resultDoc.Sections.Count);
    }

    [Theory]
    [InlineData("GET")]
    [InlineData("GeT")]
    [InlineData("get")]
    public void Execute_GetOperationIsCaseInsensitive(string operation)
    {
        var docPath = CreateWordDocument($"test_case_{operation}.docx");
        var result = _tool.Execute(operation, docPath);
        Assert.Contains("\"sections\"", result);
    }

    [Theory]
    [InlineData("INSERT")]
    [InlineData("InSeRt")]
    [InlineData("insert")]
    public void Execute_InsertOperationIsCaseInsensitive(string operation)
    {
        var docPath = CreateWordDocumentWithContent($"test_insert_{operation}.docx", "Content");
        var outputPath = CreateTestFilePath($"test_insert_{operation}_output.docx");
        var result = _tool.Execute(operation, docPath, outputPath: outputPath,
            sectionBreakType: "NextPage", insertAtParagraphIndex: 0);
        Assert.StartsWith("Section break inserted", result);
    }

    [Theory]
    [InlineData("DELETE")]
    [InlineData("DeLeTe")]
    [InlineData("delete")]
    public void Execute_DeleteOperationIsCaseInsensitive(string operation)
    {
        var docPath = CreateWordDocument($"test_delete_{operation}.docx");
        var doc = new Document(docPath);
        var builder = new DocumentBuilder(doc);
        builder.InsertBreak(BreakType.SectionBreakNewPage);
        builder.Writeln("Section 2");
        doc.Save(docPath);

        var outputPath = CreateTestFilePath($"test_delete_{operation}_output.docx");
        var result = _tool.Execute(operation, docPath, outputPath: outputPath, sectionIndex: 1);
        Assert.StartsWith("Deleted", result);
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
    public void InsertSection_WithMissingSectionBreakType_ShouldThrowArgumentException()
    {
        var docPath = CreateWordDocumentWithContent("test_missing_break_type.docx", "Test content");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("insert", docPath, sectionBreakType: "", insertAtParagraphIndex: 0));

        Assert.Contains("sectionBreakType", ex.Message);
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
    public void InsertSection_WithNegativeSectionIndex_ShouldThrowException()
    {
        var docPath = CreateWordDocumentWithContent("test_neg_sec_idx.docx", "Content");
        Assert.Throws<ArgumentException>(() =>
            _tool.Execute("insert", docPath,
                sectionBreakType: "NextPage", sectionIndex: -1, insertAtParagraphIndex: 0));
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
    public void GetSections_WithInvalidIndex_ShouldThrowException()
    {
        var docPath = CreateWordDocumentWithContent("test_get_invalid_index.docx", "Content");
        var exception = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("get", docPath, sectionIndex: 99));
        Assert.Contains("sectionIndex must be between", exception.Message);
    }

    [Fact]
    public void GetSections_WithNegativeSectionIndex_ShouldThrowException()
    {
        var docPath = CreateWordDocument("test_get_neg_idx.docx");
        Assert.Throws<ArgumentException>(() =>
            _tool.Execute("get", docPath, sectionIndex: -1));
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

    #region Session

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
        Assert.StartsWith("Section break inserted", result);

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
        Assert.StartsWith("Deleted", result);

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

        var result = _tool.Execute("get", docPath1, sessionId);

        Assert.Contains("\"sections\"", result);
        Assert.Contains("\"index\": 1", result);
    }

    #endregion
}