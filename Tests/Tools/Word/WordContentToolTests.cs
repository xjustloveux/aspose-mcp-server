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

    #region General

    [Fact]
    public void GetContent_ShouldReturnContent()
    {
        var docPath = CreateWordDocumentWithContent("test_get_content.docx", "Test content for extraction");
        var result = _tool.Execute("get_content", docPath);
        Assert.NotNull(result);
        Assert.Contains("content", result, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void GetContent_WithMaxChars_ShouldLimitContent()
    {
        var docPath = CreateWordDocumentWithContent("test_max_chars.docx", "This is a long test document content");
        var result = _tool.Execute("get_content", docPath, maxChars: 10);
        Assert.Contains("More content available", result);
    }

    [Fact]
    public void GetContent_WithOffset_ShouldStartFromOffset()
    {
        var docPath = CreateWordDocumentWithContent("test_offset.docx", "First part second part");
        var result = _tool.Execute("get_content", docPath, offset: 5);
        Assert.Contains("Showing chars 5", result);
    }

    [Fact]
    public void GetContent_WithOffsetAndMaxChars_ShouldPaginate()
    {
        var docPath = CreateWordDocumentWithContent("test_paginate.docx", "0123456789ABCDEFGHIJ");
        var result = _tool.Execute("get_content", docPath, offset: 5, maxChars: 5);
        Assert.Contains("Showing chars 5 to 10", result);
    }

    [Fact]
    public void GetContentDetailed_ShouldReturnDetailedContent()
    {
        var docPath = CreateWordDocumentWithContent("test_detailed.docx", "Detailed content");
        var result = _tool.Execute("get_content_detailed", docPath);
        Assert.Contains("Detailed Document Content", result);
        Assert.Contains("Body Content", result);
    }

    [Fact]
    public void GetContentDetailed_WithIncludeHeaders_ShouldIncludeHeaders()
    {
        var docPath = CreateWordDocumentWithContent("test_headers.docx", "Content with headers");
        var result = _tool.Execute("get_content_detailed", docPath, includeHeaders: true);
        Assert.Contains("Headers", result);
    }

    [Fact]
    public void GetContentDetailed_WithIncludeFooters_ShouldIncludeFooters()
    {
        var docPath = CreateWordDocumentWithContent("test_footers.docx", "Content with footers");
        var result = _tool.Execute("get_content_detailed", docPath, includeFooters: true);
        Assert.Contains("Footers", result);
    }

    [Fact]
    public void GetContentDetailed_WithBothHeadersAndFooters_ShouldIncludeBoth()
    {
        var docPath = CreateWordDocumentWithContent("test_both.docx", "Content");
        var result = _tool.Execute("get_content_detailed", docPath, includeHeaders: true, includeFooters: true);
        Assert.Contains("Headers", result);
        Assert.Contains("Footers", result);
    }

    [Fact]
    public void GetStatistics_ShouldReturnStatistics()
    {
        var docPath = CreateWordDocumentWithContent("test_statistics.docx", "Test document for statistics");
        var result = _tool.Execute("get_statistics", docPath);
        Assert.Contains("\"pages\"", result);
        Assert.Contains("\"words\"", result);
        Assert.Contains("\"paragraphs\"", result);
        Assert.Contains("\"characters\"", result);
        Assert.Contains("\"tables\"", result);
        Assert.Contains("\"images\"", result);
    }

    [Fact]
    public void GetStatistics_WithIncludeFootnotes_ShouldIncludeFootnotes()
    {
        var docPath = CreateWordDocumentWithContent("test_statistics_fn.docx", "Test content");
        var result = _tool.Execute("get_statistics", docPath, includeFootnotes: true);
        Assert.Contains("\"footnotes\"", result);
        Assert.Contains("\"footnotesIncluded\": true", result);
    }

    [Fact]
    public void GetStatistics_WithoutFootnotes_ShouldExcludeFootnotes()
    {
        var docPath = CreateWordDocumentWithContent("test_statistics_no_fn.docx", "Test content");
        var result = _tool.Execute("get_statistics", docPath, includeFootnotes: false);
        Assert.Contains("\"footnotesIncluded\": false", result);
    }

    [Fact]
    public void GetDocumentInfo_ShouldReturnDocumentInfo()
    {
        var docPath = CreateWordDocument("test_doc_info.docx");
        var result = _tool.Execute("get_document_info", docPath);
        Assert.Contains("\"title\"", result);
        Assert.Contains("\"author\"", result);
        Assert.Contains("\"sections\"", result);
        Assert.Contains("\"created\"", result);
        Assert.Contains("\"modified\"", result);
    }

    [Fact]
    public void GetDocumentInfo_WithIncludeTabStops_ShouldIncludeTabStops()
    {
        var docPath = CreateWordDocument("test_doc_info_tabs.docx");
        var result = _tool.Execute("get_document_info", docPath, includeTabStops: true);
        Assert.Contains("\"tabStopsIncluded\": true", result);
    }

    [Fact]
    public void GetDocumentInfo_WithoutTabStops_ShouldExcludeTabStops()
    {
        var docPath = CreateWordDocument("test_doc_info_no_tabs.docx");
        var result = _tool.Execute("get_document_info", docPath, includeTabStops: false);
        Assert.Contains("\"tabStopsIncluded\": false", result);
    }

    [Theory]
    [InlineData("GET_CONTENT")]
    [InlineData("Get_Content")]
    [InlineData("get_content")]
    public void Operation_ShouldBeCaseInsensitive_GetContent(string operation)
    {
        var docPath = CreateWordDocumentWithContent($"test_case_{operation.Replace("_", "")}.docx", "Test");
        var result = _tool.Execute(operation, docPath);
        Assert.Contains("Document Content", result);
    }

    [Theory]
    [InlineData("GET_STATISTICS")]
    [InlineData("Get_Statistics")]
    [InlineData("get_statistics")]
    public void Operation_ShouldBeCaseInsensitive_GetStatistics(string operation)
    {
        var docPath = CreateWordDocumentWithContent($"test_case_stats_{operation.Replace("_", "")}.docx", "Test");
        var result = _tool.Execute(operation, docPath);
        Assert.Contains("\"pages\"", result);
    }

    [Theory]
    [InlineData("GET_DOCUMENT_INFO")]
    [InlineData("Get_Document_Info")]
    [InlineData("get_document_info")]
    public void Operation_ShouldBeCaseInsensitive_GetDocumentInfo(string operation)
    {
        var docPath = CreateWordDocument($"test_case_info_{operation.Replace("_", "")}.docx");
        var result = _tool.Execute(operation, docPath);
        Assert.Contains("\"title\"", result);
    }

    [Theory]
    [InlineData("GET_CONTENT_DETAILED")]
    [InlineData("Get_Content_Detailed")]
    [InlineData("get_content_detailed")]
    public void Operation_ShouldBeCaseInsensitive_GetContentDetailed(string operation)
    {
        var docPath = CreateWordDocumentWithContent($"test_case_detailed_{operation.Replace("_", "")}.docx", "Test");
        var result = _tool.Execute(operation, docPath);
        Assert.Contains("Detailed Document Content", result);
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
    }

    [Fact]
    public void GetContent_WithOffsetBeyondEnd_ShouldReturnEmpty()
    {
        var docPath = CreateWordDocumentWithContent("test_offset_beyond.docx", "Short");
        var result = _tool.Execute("get_content", docPath, offset: 10000);
        Assert.NotNull(result);
        Assert.Contains("Document Content", result);
    }

    [Fact]
    public void GetContent_WithNegativeOffset_ShouldHandleGracefully()
    {
        var docPath = CreateWordDocumentWithContent("test_negative_offset.docx", "Test content");
        var result = _tool.Execute("get_content", docPath, offset: -1);
        Assert.NotNull(result);
        Assert.Contains("content", result, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void GetContent_WithZeroMaxChars_ShouldReturnEmpty()
    {
        var docPath = CreateWordDocumentWithContent("test_zero_max.docx", "Test content");
        var result = _tool.Execute("get_content", docPath, maxChars: 0);
        Assert.Contains("More content available", result);
    }

    [Fact]
    public void GetContent_WithVeryLargeMaxChars_ShouldReturnAllContent()
    {
        var docPath = CreateWordDocumentWithContent("test_large_max.docx", "Test content");
        var result = _tool.Execute("get_content", docPath, maxChars: 1000000);
        Assert.Contains("Test content", result);
        Assert.DoesNotContain("More content available", result);
    }

    #endregion

    #region Session

    [Fact]
    public void GetContent_WithSessionId_ShouldReturnContent()
    {
        var docPath = CreateWordDocumentWithContent("test_session_content.docx", "Session content for extraction");
        var sessionId = OpenSession(docPath);
        var result = _tool.Execute("get_content", sessionId: sessionId);
        Assert.Contains("Session content", result);
    }

    [Fact]
    public void GetStatistics_WithSessionId_ShouldReturnStatistics()
    {
        var docPath = CreateWordDocumentWithContent("test_session_stats.docx", "Session document for statistics");
        var sessionId = OpenSession(docPath);
        var result = _tool.Execute("get_statistics", sessionId: sessionId);
        Assert.Contains("\"pages\"", result);
        Assert.Contains("\"words\"", result);
    }

    [Fact]
    public void GetDocumentInfo_WithSessionId_ShouldReturnDocumentInfo()
    {
        var docPath = CreateWordDocument("test_session_info.docx");
        var sessionId = OpenSession(docPath);
        var result = _tool.Execute("get_document_info", sessionId: sessionId);
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
        Assert.Contains("Detailed Document Content", result);
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
        var docPath2 = CreateWordDocumentWithContent("test_session_content2.docx", "SessionDocumentContent");
        var sessionId = OpenSession(docPath2);
        var result = _tool.Execute("get_content", docPath1, sessionId);
        Assert.Contains("SessionDocumentContent", result);
        Assert.DoesNotContain("PathDocumentContent", result);
    }

    #endregion
}