using System.Text.Json;
using Aspose.Pdf;
using Aspose.Pdf.Text;
using AsposeMcpServer.Tests.Helpers;
using AsposeMcpServer.Tools.Pdf;

namespace AsposeMcpServer.Tests.Tools.Pdf;

/// <summary>
///     Integration tests for PdfTextTool.
///     Focuses on session management, file I/O, and operation routing.
///     Detailed parameter validation and business logic tests are in Handler tests.
/// </summary>
public class PdfTextToolTests : PdfTestBase
{
    private readonly PdfTextTool _tool;

    public PdfTextToolTests()
    {
        _tool = new PdfTextTool(SessionManager);
    }

    private string CreatePdfDocument(string fileName, string content = "Sample PDF Text")
    {
        var filePath = CreateTestFilePath(fileName);
        using var document = new Document();
        var page = document.Pages.Add();
        page.Paragraphs.Add(new TextFragment(content));
        document.Save(filePath);
        return filePath;
    }

    #region File I/O Smoke Tests

    [Fact]
    public void Extract_ShouldReturnJsonResult()
    {
        var pdfPath = CreatePdfDocument("test_extract.pdf");
        var result = _tool.Execute("extract", pdfPath, pageIndex: 1);
        var json = JsonSerializer.Deserialize<JsonElement>(result);
        Assert.True(json.TryGetProperty("pageIndex", out var pageIndex));
        Assert.Equal(1, pageIndex.GetInt32());
        Assert.True(json.TryGetProperty("text", out var text));
        Assert.False(string.IsNullOrEmpty(text.GetString()));
    }

    [Fact]
    public void Extract_WithIncludeFontInfo_ShouldReturnFragments()
    {
        var pdfPath = CreatePdfDocument("test_extract_font.pdf");
        var result = _tool.Execute("extract", pdfPath, pageIndex: 1, includeFontInfo: true);
        var json = JsonSerializer.Deserialize<JsonElement>(result);
        Assert.True(json.TryGetProperty("fragments", out var fragments));
        Assert.True(fragments.GetArrayLength() > 0);
    }

    [Fact]
    public void Add_ShouldAddTextToPage()
    {
        var pdfPath = CreatePdfDocument("test_add.pdf");
        var outputPath = CreateTestFilePath("test_add_output.pdf");
        var result = _tool.Execute("add", pdfPath, pageIndex: 1, outputPath: outputPath,
            text: "Added Text", x: 100, y: 700);
        Assert.True(File.Exists(outputPath));
        Assert.StartsWith("Text added to page", result);
    }

    [SkippableFact]
    public void Edit_ShouldReplaceText()
    {
        SkipInEvaluationMode(AsposeLibraryType.Pdf, "Text editing has limitations in evaluation mode");
        var pdfPath = CreatePdfDocument("test_edit.pdf");
        var outputPath = CreateTestFilePath("test_edit_output.pdf");
        var result = _tool.Execute("edit", pdfPath, pageIndex: 1, outputPath: outputPath,
            oldText: "Sample PDF Text", newText: "Updated Text", replaceAll: true);
        Assert.True(File.Exists(outputPath));
        Assert.StartsWith("Replaced", result);
    }

    #endregion

    #region Operation Routing

    [Theory]
    [InlineData("EXTRACT")]
    [InlineData("Extract")]
    [InlineData("extract")]
    public void Operation_ShouldBeCaseInsensitive(string operation)
    {
        var pdfPath = CreatePdfDocument($"test_case_{operation}.pdf");
        var result = _tool.Execute(operation, pdfPath, pageIndex: 1);
        Assert.Contains("pageIndex", result);
    }

    [Fact]
    public void Execute_WithUnknownOperation_ShouldThrowArgumentException()
    {
        var pdfPath = CreatePdfDocument("test_unknown_op.pdf");
        var ex = Assert.Throws<ArgumentException>(() => _tool.Execute("unknown", pdfPath));
        Assert.Contains("Unknown operation", ex.Message);
    }

    [Fact]
    public void Execute_WithNoPathOrSessionId_ShouldThrowException()
    {
        Assert.ThrowsAny<Exception>(() => _tool.Execute("extract"));
    }

    #endregion

    #region Session Management

    [Fact]
    public void Extract_WithSessionId_ShouldExtractFromSession()
    {
        var pdfPath = CreatePdfDocument("test_session_extract.pdf");
        var sessionId = OpenSession(pdfPath);
        var result = _tool.Execute("extract", sessionId: sessionId, pageIndex: 1);
        var json = JsonSerializer.Deserialize<JsonElement>(result);
        Assert.True(json.TryGetProperty("text", out var text));
        Assert.False(string.IsNullOrEmpty(text.GetString()));
    }

    [Fact]
    public void Add_WithSessionId_ShouldAddToSession()
    {
        var pdfPath = CreatePdfDocument("test_session_add.pdf");
        var sessionId = OpenSession(pdfPath);
        var result = _tool.Execute("add", sessionId: sessionId, pageIndex: 1,
            text: "Session Text", x: 100, y: 700);
        Assert.StartsWith("Text added to page", result);
        var document = SessionManager.GetDocument<Document>(sessionId);
        Assert.NotNull(document);
    }

    [SkippableFact]
    public void Edit_WithSessionId_ShouldEditInSession()
    {
        SkipInEvaluationMode(AsposeLibraryType.Pdf, "Text editing has limitations in evaluation mode");
        var pdfPath = CreatePdfDocument("test_session_edit.pdf");
        var sessionId = OpenSession(pdfPath);
        var result = _tool.Execute("edit", sessionId: sessionId, pageIndex: 1,
            oldText: "Sample PDF Text", newText: "Updated Session Text", replaceAll: true);
        Assert.StartsWith("Replaced", result);
    }

    [Fact]
    public void Execute_WithInvalidSessionId_ShouldThrowKeyNotFoundException()
    {
        Assert.Throws<KeyNotFoundException>(() =>
            _tool.Execute("extract", sessionId: "invalid_session", pageIndex: 1));
    }

    [SkippableFact]
    public void Execute_WithBothPathAndSessionId_ShouldPreferSessionId()
    {
        SkipInEvaluationMode(AsposeLibraryType.Pdf, "Evaluation mode replaces text content");
        var pdfPath1 = CreatePdfDocument("test_path_text.pdf", "Path Content");
        var pdfPath2 = CreatePdfDocument("test_session_text.pdf", "Session Content");
        var sessionId = OpenSession(pdfPath2);
        var result = _tool.Execute("extract", pdfPath1, sessionId, pageIndex: 1);
        var json = JsonSerializer.Deserialize<JsonElement>(result);
        var text = json.GetProperty("text").GetString();
        Assert.Contains("Session Content", text);
    }

    #endregion
}
