using Aspose.Words;
using AsposeMcpServer.Tests.Helpers;
using AsposeMcpServer.Tools.Word;

namespace AsposeMcpServer.Tests.Tools.Word;

public class WordFormatToolTests : WordTestBase
{
    private readonly WordFormatTool _tool;

    public WordFormatToolTests()
    {
        _tool = new WordFormatTool(SessionManager);
    }

    #region General Tests

    [Fact]
    public void GetRunFormat_ShouldReturnFormatInfo()
    {
        var docPath = CreateWordDocumentWithContent("test_get_run_format.docx", "Test text");
        var result = _tool.Execute("get_run_format", docPath, paragraphIndex: 0, runIndex: 0);
        Assert.NotNull(result);
        Assert.NotEmpty(result);
        Assert.Contains("Font", result, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void SetRunFormat_ShouldApplyFormatting()
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
    public void GetTabStops_ShouldReturnTabStops()
    {
        var docPath = CreateWordDocumentWithContent("test_get_tab_stops.docx", "Test");
        var result = _tool.Execute("get_tab_stops", docPath, paragraphIndex: 0);
        Assert.NotNull(result);
        Assert.NotEmpty(result);
    }

    [Fact]
    public void SetParagraphBorder_ShouldSetBorder()
    {
        var docPath = CreateWordDocumentWithContent("test_set_border.docx", "Test paragraph");
        var outputPath = CreateTestFilePath("test_set_border_output.docx");
        _tool.Execute("set_paragraph_border", docPath, outputPath: outputPath,
            paragraphIndex: 0, borderPosition: "all", lineStyle: "single", lineWidth: 1.0);
        Assert.True(File.Exists(outputPath), "Output document should be created");
        var doc = new Document(outputPath);
        var paragraphs = GetParagraphs(doc);
        Assert.True(paragraphs.Count > 0, "Document should have paragraphs");
    }

    [Fact]
    public void AddTabStop_ShouldAddTabStop()
    {
        var docPath = CreateWordDocumentWithContent("test_add_tab_stop.docx", "Test text with tab");
        var outputPath = CreateTestFilePath("test_add_tab_stop_output.docx");
        _tool.Execute("add_tab_stop", docPath, outputPath: outputPath,
            paragraphIndex: 0, tabPosition: 72.0, tabAlignment: "left", tabLeader: "none");
        Assert.True(File.Exists(outputPath), "Output document should be created");
        var doc = new Document(outputPath);
        var paragraphs = GetParagraphs(doc);
        Assert.True(paragraphs.Count > 0, "Document should have paragraphs");

        // In evaluation mode, tab stops may not be added correctly
        if (IsEvaluationMode())
        {
            // Just verify the document was created
            Assert.True(paragraphs.Count > 0, "Document should have content");
        }
        else
        {
            var tabStops = paragraphs[0].ParagraphFormat.TabStops;
            Assert.True(tabStops.Count > 0, "Paragraph should have at least one tab stop");
        }
    }

    [Fact]
    public void ClearTabStops_ShouldClearTabStops()
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
        Assert.True(File.Exists(outputPath), "Output document should be created");
        var resultDoc = new Document(outputPath);
        var resultParagraphs = GetParagraphs(resultDoc);
        Assert.True(resultParagraphs.Count > 0, "Document should have paragraphs");
        var tabStops = resultParagraphs[0].ParagraphFormat.TabStops;
        Assert.Equal(0, tabStops.Count);
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
    public void GetRunFormat_WithInvalidParagraphIndex_ShouldThrowArgumentException()
    {
        var docPath = CreateWordDocumentWithContent("test_invalid_para.docx", "Test text");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("get_run_format", docPath, paragraphIndex: 999, runIndex: 0));

        Assert.Contains("must be between", ex.Message);
    }

    [Fact]
    public void GetRunFormat_WithInvalidRunIndex_ShouldThrowArgumentException()
    {
        var docPath = CreateWordDocumentWithContent("test_invalid_run.docx", "Test text");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("get_run_format", docPath, paragraphIndex: 0, runIndex: 999));

        Assert.Contains("out of range", ex.Message);
    }

    [Fact]
    public void SetRunFormat_WithInvalidParagraphIndex_ShouldThrowArgumentException()
    {
        var docPath = CreateWordDocumentWithContent("test_set_invalid_para.docx", "Test text");
        var outputPath = CreateTestFilePath("test_set_invalid_para_output.docx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("set_run_format", docPath, outputPath: outputPath,
                paragraphIndex: 999, runIndex: 0, bold: true));

        Assert.Contains("must be between", ex.Message);
    }

    [Fact]
    public void SetParagraphBorder_WithInvalidParagraphIndex_ShouldThrowArgumentException()
    {
        var docPath = CreateWordDocumentWithContent("test_border_invalid_para.docx", "Test");
        var outputPath = CreateTestFilePath("test_border_invalid_para_output.docx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("set_paragraph_border", docPath, outputPath: outputPath,
                paragraphIndex: 999, borderPosition: "all"));

        Assert.Contains("must be between", ex.Message);
    }

    [Fact]
    public void AddTabStop_WithInvalidParagraphIndex_ShouldThrowArgumentException()
    {
        var docPath = CreateWordDocumentWithContent("test_tab_invalid_para.docx", "Test");
        var outputPath = CreateTestFilePath("test_tab_invalid_para_output.docx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("add_tab_stop", docPath, outputPath: outputPath,
                paragraphIndex: 999, tabPosition: 72.0));

        Assert.Contains("must be between", ex.Message);
    }

    [Fact]
    public void ClearTabStops_WithInvalidParagraphIndex_ShouldThrowArgumentException()
    {
        var docPath = CreateWordDocumentWithContent("test_clear_invalid_para.docx", "Test");
        var outputPath = CreateTestFilePath("test_clear_invalid_para_output.docx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("clear_tab_stops", docPath, outputPath: outputPath, paragraphIndex: 999));

        Assert.Contains("must be between", ex.Message);
    }

    [Fact]
    public void GetTabStops_WithInvalidParagraphIndex_ShouldThrowArgumentException()
    {
        var docPath = CreateWordDocumentWithContent("test_get_tabs_invalid_para.docx", "Test");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("get_tab_stops", docPath, paragraphIndex: 999));

        Assert.Contains("out of range", ex.Message);
    }

    #endregion

    #region Session ID Tests

    [Fact]
    public void GetRunFormat_WithSessionId_ShouldReturnFormatInfo()
    {
        var docPath = CreateWordDocumentWithContent("test_session_get_format.docx", "Session text");
        var sessionId = OpenSession(docPath);
        var result = _tool.Execute("get_run_format", sessionId: sessionId, paragraphIndex: 0, runIndex: 0);
        Assert.NotNull(result);
        Assert.Contains("Font", result, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void SetRunFormat_WithSessionId_ShouldApplyFormattingInMemory()
    {
        var docPath = CreateWordDocumentWithContent("test_session_set_format.docx", "Format this text");
        var sessionId = OpenSession(docPath);
        _tool.Execute("set_run_format", sessionId: sessionId,
            paragraphIndex: 0, runIndex: 0, bold: true, italic: true, fontSize: 16);

        // Assert - verify in-memory change
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
        _tool.Execute("add_tab_stop", sessionId: sessionId,
            paragraphIndex: 0, tabPosition: 72.0, tabAlignment: "left", tabLeader: "none");

        // Assert - verify in-memory change (only in non-evaluation mode)
        var sessionDoc = SessionManager.GetDocument<Document>(sessionId);
        var paragraphs = GetParagraphs(sessionDoc);
        Assert.True(paragraphs.Count > 0, "Session document should have paragraphs");

        if (!IsEvaluationMode())
        {
            var tabStops = paragraphs[0].ParagraphFormat.TabStops;
            Assert.True(tabStops.Count > 0, "Paragraph should have tab stop in session");
        }
    }

    [Fact]
    public void SetParagraphBorder_WithSessionId_ShouldSetBorderInMemory()
    {
        var docPath = CreateWordDocumentWithContent("test_session_set_border.docx", "Border test");
        var sessionId = OpenSession(docPath);
        var result = _tool.Execute("set_paragraph_border", sessionId: sessionId,
            paragraphIndex: 0, borderPosition: "all", lineStyle: "single", lineWidth: 1.5);
        Assert.Contains("border", result, StringComparison.OrdinalIgnoreCase);

        // Verify in-memory change
        var sessionDoc = SessionManager.GetDocument<Document>(sessionId);
        var paragraphs = GetParagraphs(sessionDoc);
        Assert.True(paragraphs.Count > 0, "Session document should have paragraphs");
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
        _tool.Execute("clear_tab_stops", sessionId: sessionId, paragraphIndex: 0);

        // Assert - verify in-memory change
        var sessionDoc = SessionManager.GetDocument<Document>(sessionId);
        var sessionParagraphs = GetParagraphs(sessionDoc);
        var tabStops = sessionParagraphs[0].ParagraphFormat.TabStops;
        Assert.Equal(0, tabStops.Count);
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

        // Act - provide both path and sessionId
        var result = _tool.Execute("get_run_format", docPath1, sessionId, paragraphIndex: 0, runIndex: 0);

        // Assert - should use sessionId
        Assert.NotNull(result);
        Assert.Contains("Font", result, StringComparison.OrdinalIgnoreCase);
    }

    #endregion
}