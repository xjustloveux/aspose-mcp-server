using AsposeMcpServer.Results.Word.Content;
using AsposeMcpServer.Tests.Infrastructure;
using AsposeMcpServer.Tools.Word;

namespace AsposeMcpServer.Tests.Tools.Word;

/// <summary>
///     Integration tests for WordContentTool.
///     Focuses on session management, file I/O, and operation routing.
///     Detailed parameter validation and business logic tests are in Handler tests.
/// </summary>
public class WordContentToolTests : WordTestBase
{
    private readonly WordContentTool _tool;

    public WordContentToolTests()
    {
        _tool = new WordContentTool(SessionManager);
    }

    #region File I/O Smoke Tests

    [Fact]
    public void GetContent_ShouldReturnContent()
    {
        var docPath = CreateWordDocumentWithContent("test_get_content.docx", "Test content for extraction");
        var result = _tool.Execute("get_content", docPath);
        Assert.NotNull(result);
        var data = GetResultData<GetWordContentResult>(result);
        Assert.Contains("Test content for extraction", data.Content);
    }

    [Fact]
    public void GetContentDetailed_ShouldReturnDetailedContent()
    {
        var docPath = CreateWordDocumentWithContent("test_detailed.docx", "Detailed content");
        var result = _tool.Execute("get_content_detailed", docPath);
        var data = GetResultData<GetWordContentDetailedResult>(result);
        Assert.Contains("Detailed content", data.Content);
    }

    [Fact]
    public void GetStatistics_ShouldReturnStatistics()
    {
        var docPath = CreateWordDocumentWithContent("test_statistics.docx", "Test document for statistics");
        var result = _tool.Execute("get_statistics", docPath);
        var data = GetResultData<GetWordStatisticsResult>(result);
        Assert.True(data.Pages >= 0);
        Assert.True(data.Words >= 0);
        Assert.True(data.Paragraphs >= 0);
    }

    [Fact]
    public void GetDocumentInfo_ShouldReturnDocumentInfo()
    {
        var docPath = CreateWordDocument("test_doc_info.docx");
        var result = _tool.Execute("get_document_info", docPath);
        var data = GetResultData<GetWordDocumentInfoResult>(result);
        Assert.NotNull(data.Created);
        Assert.NotNull(data.Modified);
        Assert.True(data.Sections >= 0);
    }

    #endregion

    #region Operation Routing

    [Theory]
    [InlineData("GET_CONTENT")]
    [InlineData("Get_Content")]
    [InlineData("get_content")]
    public void Operation_ShouldBeCaseInsensitive(string operation)
    {
        var docPath = CreateWordDocumentWithContent($"test_case_{operation.Replace("_", "")}.docx", "Test");
        var result = _tool.Execute(operation, docPath);
        var data = GetResultData<GetWordContentResult>(result);
        Assert.Contains("Test", data.Content);
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
    public void GetContent_WithSessionId_ShouldReturnContent()
    {
        var docPath = CreateWordDocumentWithContent("test_session_content.docx", "Session content for extraction");
        var sessionId = OpenSession(docPath);
        var result = _tool.Execute("get_content", sessionId: sessionId);
        var data = GetResultData<GetWordContentResult>(result);
        Assert.Contains("Session content", data.Content);
    }

    [Fact]
    public void GetStatistics_WithSessionId_ShouldReturnStatistics()
    {
        var docPath = CreateWordDocumentWithContent("test_session_stats.docx", "Session document for statistics");
        var sessionId = OpenSession(docPath);
        var result = _tool.Execute("get_statistics", sessionId: sessionId);
        var data = GetResultData<GetWordStatisticsResult>(result);
        Assert.True(data.Pages >= 0);
        Assert.True(data.Words >= 0);
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
        var data = GetResultData<GetWordContentResult>(result);
        Assert.Contains("SessionDocumentContent", data.Content);
        Assert.DoesNotContain("PathDocumentContent", data.Content);
    }

    #endregion
}
