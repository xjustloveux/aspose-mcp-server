using Aspose.Words;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Results.Word.Format;
using AsposeMcpServer.Tests.Infrastructure;
using AsposeMcpServer.Tools.Word;

namespace AsposeMcpServer.Tests.Tools.Word;

/// <summary>
///     Integration tests for WordFormatTool.
///     Focuses on session management, file I/O, and operation routing.
///     Detailed parameter validation and business logic tests are in Handler tests.
/// </summary>
public class WordFormatToolTests : WordTestBase
{
    private readonly WordFormatTool _tool;

    public WordFormatToolTests()
    {
        _tool = new WordFormatTool(SessionManager);
    }

    #region File I/O Smoke Tests

    [Fact]
    public void GetRunFormat_ShouldReturnFormatInfoFromFile()
    {
        var docPath = CreateWordDocumentWithContent("test_get_run_format.docx", "Test text");
        var result = _tool.Execute("get_run_format", docPath, paragraphIndex: 0, runIndex: 0);
        var data = GetResultData<GetRunFormatWordResult>(result);
        Assert.Equal(0, data.ParagraphIndex);
        Assert.Equal(0, data.RunIndex);
        Assert.NotNull(data.FontName);
    }

    [Fact]
    public void SetRunFormat_ShouldApplyFormattingAndPersistToFile()
    {
        var docPath = CreateWordDocumentWithContent("test_set_run_format.docx", "Format this");
        var outputPath = CreateTestFilePath("test_set_run_format_output.docx");
        _tool.Execute("set_run_format", docPath, outputPath: outputPath,
            paragraphIndex: 0, runIndex: 0, bold: true, fontSize: 14);
        var doc = new Document(outputPath);
        var runs = doc.GetChildNodes(NodeType.Run, true).Cast<Run>().ToList();
        if (runs.Count > 0)
        {
            Assert.True(runs[0].Font.Bold);
            Assert.Equal(14, runs[0].Font.Size);
        }
    }

    [Fact]
    public void GetTabStops_ShouldReturnTabStopsFromFile()
    {
        var docPath = CreateWordDocumentWithContent("test_get_tab_stops.docx", "Test");
        var result = _tool.Execute("get_tab_stops", docPath, paragraphIndex: 0);
        var data = GetResultData<GetTabStopsWordResult>(result);
        Assert.NotNull(data.TabStops);
        Assert.Equal(0, data.SectionIndex);
    }

    [Fact]
    public void SetParagraphBorder_ShouldSetBorderAndPersistToFile()
    {
        var docPath = CreateWordDocumentWithContent("test_set_border.docx", "Test paragraph");
        var outputPath = CreateTestFilePath("test_set_border_output.docx");
        _tool.Execute("set_paragraph_border", docPath, outputPath: outputPath,
            paragraphIndex: 0, borderPosition: "all", lineStyle: "single", lineWidth: 1.0);
        Assert.True(File.Exists(outputPath));
        var doc = new Document(outputPath);
        Assert.True(GetParagraphs(doc).Count > 0);
    }

    [Fact]
    public void AddTabStop_ShouldAddTabStopAndPersistToFile()
    {
        var docPath = CreateWordDocumentWithContent("test_add_tab_stop.docx", "Test text with tab");
        var outputPath = CreateTestFilePath("test_add_tab_stop_output.docx");
        _tool.Execute("add_tab_stop", docPath, outputPath: outputPath,
            paragraphIndex: 0, tabPosition: 72.0, tabAlignment: "left", tabLeader: "none");
        Assert.True(File.Exists(outputPath));
    }

    [Fact]
    public void ClearTabStops_ShouldClearTabStopsAndPersistToFile()
    {
        var docPath = CreateWordDocumentWithContent("test_clear_tab_stops.docx", "Test text");
        var doc = new Document(docPath);
        var paragraphs = GetParagraphs(doc);
        if (paragraphs.Count > 0)
        {
            paragraphs[0].ParagraphFormat.TabStops.Add(72.0, TabAlignment.Left, TabLeader.None);
            doc.Save(docPath);
        }

        var outputPath = CreateTestFilePath("test_clear_tab_stops_output.docx");
        _tool.Execute("clear_tab_stops", docPath, outputPath: outputPath, paragraphIndex: 0);
        var resultDoc = new Document(outputPath);
        var resultParagraphs = GetParagraphs(resultDoc);
        Assert.Equal(0, resultParagraphs[0].ParagraphFormat.TabStops.Count);
    }

    #endregion

    #region Operation Routing

    [Theory]
    [InlineData("GET_RUN_FORMAT")]
    [InlineData("Get_Run_Format")]
    [InlineData("get_run_format")]
    public void Operation_ShouldBeCaseInsensitive(string operation)
    {
        var docPath = CreateWordDocumentWithContent($"test_case_{operation.Replace("_", "")}.docx", "Test");
        var result = _tool.Execute(operation, docPath, paragraphIndex: 0, runIndex: 0);
        var data = GetResultData<GetRunFormatWordResult>(result);
        Assert.NotNull(data.FontName);
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
        Assert.ThrowsAny<Exception>(() => _tool.Execute("get_run_format", paragraphIndex: 0, runIndex: 0));
    }

    #endregion

    #region Session Management

    [Fact]
    public void GetRunFormat_WithSessionId_ShouldReturnFormatInfo()
    {
        var docPath = CreateWordDocumentWithContent("test_session_get_format.docx", "Session text");
        var sessionId = OpenSession(docPath);
        var result = _tool.Execute("get_run_format", sessionId: sessionId, paragraphIndex: 0, runIndex: 0);
        var data = GetResultData<GetRunFormatWordResult>(result);
        Assert.NotNull(data.FontName);
        var output = GetResultOutput<GetRunFormatWordResult>(result);
        Assert.Equal(sessionId, output.SessionId);
    }

    [Fact]
    public void SetRunFormat_WithSessionId_ShouldApplyFormattingInMemory()
    {
        var docPath = CreateWordDocumentWithContent("test_session_set_format.docx", "Format this text");
        var sessionId = OpenSession(docPath);
        var result = _tool.Execute("set_run_format", sessionId: sessionId,
            paragraphIndex: 0, runIndex: 0, bold: true, italic: true, fontSize: 16);
        var output = GetResultOutput<SuccessResult>(result);
        Assert.Equal(sessionId, output.SessionId);
        var sessionDoc = SessionManager.GetDocument<Document>(sessionId);
        var runs = sessionDoc.GetChildNodes(NodeType.Run, true).Cast<Run>().ToList();
        if (runs.Count > 0)
        {
            Assert.True(runs[0].Font.Bold);
            Assert.True(runs[0].Font.Italic);
            Assert.Equal(16, runs[0].Font.Size);
        }
    }

    [Fact]
    public void AddTabStop_WithSessionId_ShouldAddTabStopInMemory()
    {
        var docPath = CreateWordDocumentWithContent("test_session_add_tab.docx", "Tab test");
        var sessionId = OpenSession(docPath);
        var result = _tool.Execute("add_tab_stop", sessionId: sessionId,
            paragraphIndex: 0, tabPosition: 72.0, tabAlignment: "left", tabLeader: "none");
        var output = GetResultOutput<SuccessResult>(result);
        Assert.Equal(sessionId, output.SessionId);
        var sessionDoc = SessionManager.GetDocument<Document>(sessionId);
        var paragraphs = GetParagraphs(sessionDoc);
        Assert.True(paragraphs.Count > 0);
        Assert.True(paragraphs[0].ParagraphFormat.TabStops.Count > 0);
    }

    [Fact]
    public void SetParagraphBorder_WithSessionId_ShouldSetBorderInMemory()
    {
        var docPath = CreateWordDocumentWithContent("test_session_set_border.docx", "Border test");
        var sessionId = OpenSession(docPath);
        var result = _tool.Execute("set_paragraph_border", sessionId: sessionId,
            paragraphIndex: 0, borderPosition: "all", lineStyle: "single", lineWidth: 1.5);
        var output = GetResultOutput<SuccessResult>(result);
        Assert.Equal(sessionId, output.SessionId);
        var sessionDoc = SessionManager.GetDocument<Document>(sessionId);
        var paragraphs = GetParagraphs(sessionDoc);
        Assert.True(paragraphs.Count > 0);
    }

    [Fact]
    public void ClearTabStops_WithSessionId_ShouldClearTabStopsInMemory()
    {
        var docPath = CreateWordDocumentWithContent("test_session_clear_tabs.docx", "Tab clear test");
        var doc = new Document(docPath);
        var paragraphs = GetParagraphs(doc);
        if (paragraphs.Count > 0)
        {
            paragraphs[0].ParagraphFormat.TabStops.Add(72.0, TabAlignment.Left, TabLeader.None);
            doc.Save(docPath);
        }

        var sessionId = OpenSession(docPath);
        var result = _tool.Execute("clear_tab_stops", sessionId: sessionId, paragraphIndex: 0);
        var output = GetResultOutput<SuccessResult>(result);
        Assert.Equal(sessionId, output.SessionId);
        var sessionDoc = SessionManager.GetDocument<Document>(sessionId);
        var sessionParagraphs = GetParagraphs(sessionDoc);
        Assert.Equal(0, sessionParagraphs[0].ParagraphFormat.TabStops.Count);
    }

    [Fact]
    public void GetTabStops_WithSessionId_ShouldReturnTabStops()
    {
        var docPath = CreateWordDocumentWithContent("test_session_get_tabs.docx", "Tab test");
        var sessionId = OpenSession(docPath);
        var result = _tool.Execute("get_tab_stops", sessionId: sessionId, paragraphIndex: 0);
        var data = GetResultData<GetTabStopsWordResult>(result);
        Assert.NotNull(data.TabStops);
        var output = GetResultOutput<GetTabStopsWordResult>(result);
        Assert.Equal(sessionId, output.SessionId);
    }

    [Fact]
    public void Execute_WithInvalidSessionId_ShouldThrowKeyNotFoundException()
    {
        Assert.Throws<KeyNotFoundException>(() =>
            _tool.Execute("get_run_format", sessionId: "invalid_session_id", paragraphIndex: 0, runIndex: 0));
    }

    [Fact]
    public void Execute_WithBothPathAndSessionId_ShouldPreferSessionId()
    {
        var docPath1 = CreateWordDocumentWithContent("test_path_format.docx", "PathContent");
        var docPath2 = CreateWordDocumentWithContent("test_session_format.docx", "SessionContent");
        var sessionId = OpenSession(docPath2);
        var result = _tool.Execute("get_run_format", docPath1, sessionId, paragraphIndex: 0, runIndex: 0);
        var data = GetResultData<GetRunFormatWordResult>(result);
        Assert.NotNull(data.FontName);
        var output = GetResultOutput<GetRunFormatWordResult>(result);
        Assert.Equal(sessionId, output.SessionId);
    }

    #endregion
}
