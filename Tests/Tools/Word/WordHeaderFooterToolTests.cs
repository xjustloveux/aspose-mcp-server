using Aspose.Words;
using AsposeMcpServer.Results;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Results.Word.HeaderFooter;
using AsposeMcpServer.Tests.Infrastructure;
using AsposeMcpServer.Tools.Word;

namespace AsposeMcpServer.Tests.Tools.Word;

/// <summary>
///     Integration tests for WordHeaderFooterTool.
///     Focuses on session management, file I/O, and operation routing.
///     Detailed parameter validation and business logic tests are in Handler tests.
/// </summary>
public class WordHeaderFooterToolTests : WordTestBase
{
    private readonly WordHeaderFooterTool _tool;

    public WordHeaderFooterToolTests()
    {
        _tool = new WordHeaderFooterTool(SessionManager);
    }

    #region File I/O Smoke Tests

    [Fact]
    public void SetHeaderText_ShouldSetHeaderTextAndPersistToFile()
    {
        var docPath = CreateWordDocument("test_set_header_text.docx");
        var outputPath = CreateTestFilePath("test_set_header_text_output.docx");
        _tool.Execute("set_header_text", docPath, outputPath: outputPath,
            headerLeft: "Left Header", headerCenter: "Center Header");
        var doc = new Document(outputPath);
        var header = doc.FirstSection.HeadersFooters[HeaderFooterType.HeaderPrimary];
        Assert.NotNull(header);
        Assert.Contains("Left", header.GetText(), StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void SetFooterText_ShouldSetFooterTextAndPersistToFile()
    {
        var docPath = CreateWordDocument("test_set_footer_text.docx");
        var outputPath = CreateTestFilePath("test_set_footer_text_output.docx");
        _tool.Execute("set_footer_text", docPath, outputPath: outputPath,
            footerLeft: "Page", footerRight: "{PAGE}");
        var doc = new Document(outputPath);
        var footer = doc.FirstSection.HeadersFooters[HeaderFooterType.FooterPrimary];
        Assert.NotNull(footer);
        Assert.Contains("Page", footer.GetText(), StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void GetHeadersFooters_ShouldReturnHeadersFootersFromFile()
    {
        var docPath = CreateWordDocument("test_get_headers_footers.docx");
        var doc = new Document(docPath);
        var header = doc.FirstSection.HeadersFooters[HeaderFooterType.HeaderPrimary];
        if (header == null)
        {
            header = new HeaderFooter(doc, HeaderFooterType.HeaderPrimary);
            doc.FirstSection.HeadersFooters.Add(header);
        }

        header.AppendParagraph("Test Header");
        doc.Save(docPath);

        var result = _tool.Execute("get", docPath);
        var data = GetResultData<GetHeadersFootersResult>(result);
        Assert.True(data.TotalSections > 0);
    }

    [Fact]
    public void SetHeaderLine_ShouldSetHeaderLineAndPersistToFile()
    {
        var docPath = CreateWordDocument("test_set_header_line.docx");
        var outputPath = CreateTestFilePath("test_set_header_line_output.docx");
        _tool.Execute("set_header_line", docPath, outputPath: outputPath, lineStyle: "single");
        Assert.True(File.Exists(outputPath));
        var doc = new Document(outputPath);
        Assert.NotNull(doc.FirstSection.HeadersFooters[HeaderFooterType.HeaderPrimary]);
    }

    [Fact]
    public void SetFooterLine_ShouldSetFooterLineAndPersistToFile()
    {
        var docPath = CreateWordDocument("test_set_footer_line.docx");
        var outputPath = CreateTestFilePath("test_set_footer_line_output.docx");
        _tool.Execute("set_footer_line", docPath, outputPath: outputPath, lineStyle: "single");
        Assert.True(File.Exists(outputPath));
        var doc = new Document(outputPath);
        Assert.NotNull(doc.FirstSection.HeadersFooters[HeaderFooterType.FooterPrimary]);
    }

    [Fact]
    public void SetHeaderFooter_ShouldSetBothAndPersistToFile()
    {
        var docPath = CreateWordDocument("test_set_header_footer.docx");
        var outputPath = CreateTestFilePath("test_set_header_footer_output.docx");
        _tool.Execute("set_header_footer", docPath, outputPath: outputPath,
            headerLeft: "Left Header", footerCenter: "Center Footer");
        Assert.True(File.Exists(outputPath));
        var doc = new Document(outputPath);
        Assert.NotNull(doc.FirstSection.HeadersFooters[HeaderFooterType.HeaderPrimary]);
        Assert.NotNull(doc.FirstSection.HeadersFooters[HeaderFooterType.FooterPrimary]);
    }

    #endregion

    #region Operation Routing

    [Theory]
    [InlineData("SET_HEADER_TEXT")]
    [InlineData("Set_Header_Text")]
    [InlineData("set_header_text")]
    public void Operation_ShouldBeCaseInsensitive(string operation)
    {
        var docPath = CreateWordDocument($"test_case_{operation.Replace("_", "")}.docx");
        var outputPath = CreateTestFilePath($"test_case_{operation.Replace("_", "")}_output.docx");
        var result = _tool.Execute(operation, docPath, outputPath: outputPath, headerLeft: "Test");
        Assert.IsType<FinalizedResult<SuccessResult>>(result);
        var doc = new Document(outputPath);
        var header = doc.FirstSection.HeadersFooters[HeaderFooterType.HeaderPrimary];
        Assert.NotNull(header);
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
        Assert.ThrowsAny<Exception>(() => _tool.Execute("get"));
    }

    #endregion

    #region Session Management

    [Fact]
    public void SetHeaderText_WithSessionId_ShouldSetHeaderInMemory()
    {
        var docPath = CreateWordDocument("test_session_set_header.docx");
        var sessionId = OpenSession(docPath);
        var result = _tool.Execute("set_header_text", sessionId: sessionId,
            headerLeft: "Session Left", headerCenter: "Session Center");
        Assert.IsType<FinalizedResult<SuccessResult>>(result);

        var doc = SessionManager.GetDocument<Document>(sessionId);
        var header = doc.FirstSection.HeadersFooters[HeaderFooterType.HeaderPrimary];
        Assert.NotNull(header);
        Assert.Contains("Session", header.GetText(), StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void SetFooterText_WithSessionId_ShouldSetFooterInMemory()
    {
        var docPath = CreateWordDocument("test_session_set_footer.docx");
        var sessionId = OpenSession(docPath);
        var result = _tool.Execute("set_footer_text", sessionId: sessionId,
            footerLeft: "Session Footer", footerRight: "{PAGE}");
        Assert.IsType<FinalizedResult<SuccessResult>>(result);

        var doc = SessionManager.GetDocument<Document>(sessionId);
        var footer = doc.FirstSection.HeadersFooters[HeaderFooterType.FooterPrimary];
        Assert.NotNull(footer);
        if (!IsEvaluationMode())
            Assert.Contains("Session Footer", footer.GetText());
    }

    [SkippableFact]
    public void GetHeadersFooters_WithSessionId_ShouldReturnHeadersFooters()
    {
        SkipInEvaluationMode(AsposeLibraryType.Words, "Evaluation mode blocks document structure modification");
        var docPath = CreateWordDocument("test_session_get_hf.docx");
        var doc = new Document(docPath);
        var header = new HeaderFooter(doc, HeaderFooterType.HeaderPrimary);
        doc.FirstSection.HeadersFooters.Add(header);
        header.AppendParagraph("Session Header");
        doc.Save(docPath);

        var sessionId = OpenSession(docPath);
        var result = _tool.Execute("get", sessionId: sessionId);
        var data = GetResultData<GetHeadersFootersResult>(result);
        Assert.True(data.TotalSections > 0);
        var output = GetResultOutput<GetHeadersFootersResult>(result);
        Assert.True(output.IsSession);
    }

    [Fact]
    public void SetHeaderLine_WithSessionId_ShouldSetLineInMemory()
    {
        var docPath = CreateWordDocument("test_session_header_line.docx");
        var sessionId = OpenSession(docPath);
        var result = _tool.Execute("set_header_line", sessionId: sessionId, lineStyle: "double");
        Assert.IsType<FinalizedResult<SuccessResult>>(result);

        var doc = SessionManager.GetDocument<Document>(sessionId);
        Assert.NotNull(doc.FirstSection.HeadersFooters[HeaderFooterType.HeaderPrimary]);
    }

    [Fact]
    public void SetFooterLine_WithSessionId_ShouldSetLineInMemory()
    {
        var docPath = CreateWordDocument("test_session_footer_line.docx");
        var sessionId = OpenSession(docPath);
        var result = _tool.Execute("set_footer_line", sessionId: sessionId, lineStyle: "thick");
        Assert.IsType<FinalizedResult<SuccessResult>>(result);

        var doc = SessionManager.GetDocument<Document>(sessionId);
        Assert.NotNull(doc.FirstSection.HeadersFooters[HeaderFooterType.FooterPrimary]);
    }

    [Fact]
    public void SetHeaderFooter_WithSessionId_ShouldSetBothInMemory()
    {
        var docPath = CreateWordDocument("test_session_hf_both.docx");
        var sessionId = OpenSession(docPath);
        var result = _tool.Execute("set_header_footer", sessionId: sessionId,
            headerCenter: "Session Header", footerCenter: "Session Footer");
        Assert.IsType<FinalizedResult<SuccessResult>>(result);

        var doc = SessionManager.GetDocument<Document>(sessionId);
        Assert.NotNull(doc.FirstSection.HeadersFooters[HeaderFooterType.HeaderPrimary]);
        Assert.NotNull(doc.FirstSection.HeadersFooters[HeaderFooterType.FooterPrimary]);
    }

    [Fact]
    public void Execute_WithInvalidSessionId_ShouldThrowKeyNotFoundException()
    {
        Assert.Throws<KeyNotFoundException>(() =>
            _tool.Execute("get", sessionId: "invalid_session_id"));
    }

    [SkippableFact]
    public void Execute_WithBothPathAndSessionId_ShouldPreferSessionId()
    {
        SkipInEvaluationMode(AsposeLibraryType.Words, "Evaluation mode blocks document structure modification");
        var docPath1 = CreateWordDocument("test_path_hf.docx");
        var doc1 = new Document(docPath1);
        var header1 = new HeaderFooter(doc1, HeaderFooterType.HeaderPrimary);
        doc1.FirstSection.HeadersFooters.Add(header1);
        header1.AppendParagraph("Path Header");
        doc1.Save(docPath1);

        var docPath2 = CreateWordDocument("test_session_hf.docx");
        var doc2 = new Document(docPath2);
        var header2 = new HeaderFooter(doc2, HeaderFooterType.HeaderPrimary);
        doc2.FirstSection.HeadersFooters.Add(header2);
        header2.AppendParagraph("Session Header Unique");
        doc2.Save(docPath2);

        var sessionId = OpenSession(docPath2);
        var result = _tool.Execute("get", docPath1, sessionId);
        var data = GetResultData<GetHeadersFootersResult>(result);
        Assert.True(data.TotalSections > 0);
        var output = GetResultOutput<GetHeadersFootersResult>(result);
        Assert.True(output.IsSession);
    }

    #endregion
}
