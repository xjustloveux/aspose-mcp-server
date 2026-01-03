using AsposeMcpServer.Tests.Helpers;
using AsposeMcpServer.Tools.Word;

namespace AsposeMcpServer.Tests.Tools.Word;

public class WordContentToolTests : WordTestBase
{
    private readonly WordContentTool _tool;

    public WordContentToolTests()
    {
        _tool = new WordContentTool(SessionManager);
    }

    #region General Tests

    [Fact]
    public void GetContent_ShouldReturnContent()
    {
        var docPath = CreateWordDocumentWithContent("test_get_content.docx", "Test content for extraction");
        var result = _tool.Execute("get_content", docPath);
        Assert.NotNull(result);
        Assert.NotEmpty(result);
        Assert.Contains("content", result, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void GetContentDetailed_ShouldReturnDetailedContent()
    {
        var docPath = CreateWordDocumentWithContent("test_get_content_detailed.docx", "Detailed content");
        var result = _tool.Execute("get_content_detailed", docPath,
            includeHeaders: true, includeFooters: true);
        Assert.NotNull(result);
        Assert.NotEmpty(result);
    }

    [Fact]
    public void GetStatistics_ShouldReturnStatistics()
    {
        var docPath = CreateWordDocumentWithContent("test_get_statistics.docx", "Test document for statistics");
        var result = _tool.Execute("get_statistics", docPath);
        Assert.NotNull(result);
        Assert.NotEmpty(result);
        // Check for JSON properties instead of plain text header
        Assert.Contains("\"pages\"", result);
        Assert.Contains("\"words\"", result);
        Assert.Contains("\"paragraphs\"", result);
    }

    [Fact]
    public void GetDocumentInfo_ShouldReturnDocumentInfo()
    {
        var docPath = CreateWordDocument("test_get_document_info.docx");
        var result = _tool.Execute("get_document_info", docPath);
        Assert.NotNull(result);
        Assert.NotEmpty(result);
        // Check for JSON properties instead of plain text header
        Assert.Contains("\"title\"", result);
        Assert.Contains("\"author\"", result);
        Assert.Contains("\"sections\"", result);
    }

    [Fact]
    public void GetStatistics_WithIncludeFootnotes_ShouldRespectParameter()
    {
        var docPath = CreateWordDocumentWithContent("test_statistics_footnotes.docx", "Test content");
        var result = _tool.Execute("get_statistics", docPath, includeFootnotes: false);
        Assert.NotNull(result);
        Assert.NotEmpty(result);
        // Check for JSON property indicating footnotes are not included
        Assert.Contains("\"footnotesIncluded\": false", result);
    }

    [Fact]
    public void GetDocumentInfo_WithIncludeTabStops_ShouldIncludeTabStops()
    {
        var docPath = CreateWordDocument("test_document_info_tabs.docx");
        var result = _tool.Execute("get_document_info", docPath, includeTabStops: true);
        Assert.NotNull(result);
        Assert.NotEmpty(result);
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
    }

    [Fact]
    public void GetContent_WithOffsetBeyondEnd_ShouldReturnEmpty()
    {
        var docPath = CreateWordDocumentWithContent("test_offset_beyond.docx", "Short");
        var result = _tool.Execute("get_content", docPath, offset: 10000);

        // Assert - should not throw, just return empty content section
        Assert.NotNull(result);
        Assert.Contains("Document Content", result);
    }

    [Fact]
    public void GetContent_WithMaxChars_ShouldLimitContent()
    {
        var docPath = CreateWordDocumentWithContent("test_max_chars.docx", "This is a long test document content");
        var result = _tool.Execute("get_content", docPath, maxChars: 10);
        Assert.NotNull(result);
        Assert.Contains("More content available", result);
    }

    [Fact]
    public void GetContent_WithNegativeOffset_ShouldHandleGracefully()
    {
        var docPath = CreateWordDocumentWithContent("test_negative_offset.docx", "Test content");
        var result = _tool.Execute("get_content", docPath, offset: -1);

        // Assert - negative offset should be treated as 0
        Assert.NotNull(result);
        Assert.Contains("content", result, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void GetContent_WithZeroMaxChars_ShouldReturnEmpty()
    {
        var docPath = CreateWordDocumentWithContent("test_zero_max_chars.docx", "Test content");
        var result = _tool.Execute("get_content", docPath, maxChars: 0);
        Assert.NotNull(result);
        Assert.Contains("More content available", result);
    }

    #endregion

    #region Session ID Tests

    [Fact]
    public void GetContent_WithSessionId_ShouldReturnContent()
    {
        var docPath = CreateWordDocumentWithContent("test_session_get_content.docx", "Session content for extraction");
        var sessionId = OpenSession(docPath);
        var result = _tool.Execute("get_content", sessionId: sessionId);
        Assert.NotNull(result);
        Assert.Contains("Session content", result);
    }

    [Fact]
    public void GetStatistics_WithSessionId_ShouldReturnStatistics()
    {
        var docPath =
            CreateWordDocumentWithContent("test_session_statistics.docx", "Session document for statistics testing");
        var sessionId = OpenSession(docPath);
        var result = _tool.Execute("get_statistics", sessionId: sessionId);
        Assert.NotNull(result);
        Assert.Contains("\"pages\"", result);
        Assert.Contains("\"words\"", result);
    }

    [Fact]
    public void GetDocumentInfo_WithSessionId_ShouldReturnDocumentInfo()
    {
        var docPath = CreateWordDocument("test_session_document_info.docx");
        var sessionId = OpenSession(docPath);
        var result = _tool.Execute("get_document_info", sessionId: sessionId);
        Assert.NotNull(result);
        Assert.Contains("\"title\"", result);
        Assert.Contains("\"sections\"", result);
    }

    [Fact]
    public void GetContentDetailed_WithSessionId_ShouldReturnDetailedContent()
    {
        var docPath = CreateWordDocumentWithContent("test_session_detailed.docx", "Session detailed content");
        var sessionId = OpenSession(docPath);
        var result = _tool.Execute("get_content_detailed", sessionId: sessionId,
            includeHeaders: true, includeFooters: true);
        Assert.NotNull(result);
        Assert.NotEmpty(result);
    }

    [Fact]
    public void Execute_WithInvalidSessionId_ShouldThrowKeyNotFoundException()
    {
        Assert.Throws<KeyNotFoundException>(() =>
            _tool.Execute("get_content", sessionId: "invalid_session_id"));
    }

    [Fact]
    public void Execute_WithBothPathAndSessionId_ShouldPreferSessionId()
    {
        var docPath1 = CreateWordDocumentWithContent("test_path_content.docx", "PathDocumentContent");
        var docPath2 = CreateWordDocumentWithContent("test_session_content.docx", "SessionDocumentContent");

        var sessionId = OpenSession(docPath2);

        // Act - provide both path and sessionId
        var result = _tool.Execute("get_content", docPath1, sessionId);

        // Assert - should use sessionId, returning SessionDocumentContent not PathDocumentContent
        Assert.Contains("SessionDocumentContent", result);
        Assert.DoesNotContain("PathDocumentContent", result);
    }

    #endregion
}