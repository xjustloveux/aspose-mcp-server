using Aspose.Words;
using AsposeMcpServer.Tests.Helpers;
using AsposeMcpServer.Tools.Word;

namespace AsposeMcpServer.Tests.Tools.Word;

/// <summary>
///     Integration tests for WordSectionTool.
///     Focuses on session management, file I/O, and operation routing.
///     Detailed parameter validation and business logic tests are in Handler tests.
/// </summary>
public class WordSectionToolTests : WordTestBase
{
    private readonly WordSectionTool _tool;

    public WordSectionToolTests()
    {
        _tool = new WordSectionTool(SessionManager);
    }

    #region File I/O Smoke Tests

    [Fact]
    public void InsertSection_ShouldInsertSection()
    {
        var docPath = CreateWordDocumentWithContent("test_insert_section.docx", "Content before section");
        var outputPath = CreateTestFilePath("test_insert_section_output.docx");
        var result = _tool.Execute("insert", docPath, outputPath: outputPath,
            sectionBreakType: "NextPage", insertAtParagraphIndex: 0);
        Assert.StartsWith("Section break inserted", result);
        var doc = new Document(outputPath);
        Assert.True(doc.Sections.Count > 1);
    }

    [Fact]
    public void GetSections_ShouldReturnSectionsInfo()
    {
        var docPath = CreateWordDocument("test_get_sections.docx");
        var doc = new Document(docPath);
        var builder = new DocumentBuilder(doc);
        builder.InsertBreak(BreakType.SectionBreakNewPage);
        doc.Save(docPath);
        var result = _tool.Execute("get", docPath);
        Assert.NotNull(result);
        Assert.Contains("\"sections\"", result);
        Assert.Contains("\"sectionBreak\"", result);
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

        var outputPath = CreateTestFilePath("test_delete_section_output.docx");
        var result = _tool.Execute("delete", docPath, outputPath: outputPath, sectionIndex: 1);
        var resultDoc = new Document(outputPath);
        Assert.Equal(sectionsBefore - 1, resultDoc.Sections.Count);
        Assert.StartsWith("Deleted", result);
    }

    #endregion

    #region Operation Routing

    [Theory]
    [InlineData("GET")]
    [InlineData("GeT")]
    [InlineData("get")]
    public void Operation_ShouldBeCaseInsensitive(string operation)
    {
        var docPath = CreateWordDocument($"test_case_{operation}.docx");
        var result = _tool.Execute(operation, docPath);
        Assert.Contains("\"sections\"", result);
    }

    [Fact]
    public void Execute_WithUnknownOperation_ShouldThrowArgumentException()
    {
        var docPath = CreateWordDocumentWithContent("test_unknown_op.docx", "Test content");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("unknown_operation", docPath));
        Assert.Contains("Unknown operation", ex.Message);
    }

    #endregion

    #region Session Management

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
