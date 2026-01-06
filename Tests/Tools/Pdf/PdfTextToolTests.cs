using System.Text.Json;
using Aspose.Pdf;
using Aspose.Pdf.Text;
using AsposeMcpServer.Tests.Helpers;
using AsposeMcpServer.Tools.Pdf;

namespace AsposeMcpServer.Tests.Tools.Pdf;

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

    private string CreateMultiPagePdf(string fileName, int pageCount)
    {
        var filePath = CreateTestFilePath(fileName);
        using var document = new Document();
        for (var i = 1; i <= pageCount; i++)
        {
            var page = document.Pages.Add();
            page.Paragraphs.Add(new TextFragment($"Page {i} Content"));
        }

        document.Save(filePath);
        return filePath;
    }

    #region General

    [Fact]
    public void Extract_ShouldReturnJsonResult()
    {
        var pdfPath = CreatePdfDocument("test_extract.pdf");
        var result = _tool.Execute("extract", pdfPath, pageIndex: 1);
        var json = JsonSerializer.Deserialize<JsonElement>(result);
        Assert.True(json.TryGetProperty("pageIndex", out var pageIndex));
        Assert.Equal(1, pageIndex.GetInt32());
        Assert.True(json.TryGetProperty("totalPages", out _));
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
        Assert.True(json.TryGetProperty("fragmentCount", out var count));
        Assert.True(count.GetInt32() > 0);
    }

    [Fact]
    public void Extract_WithRawMode_ShouldExtractRawText()
    {
        var pdfPath = CreatePdfDocument("test_extract_raw.pdf");
        var result = _tool.Execute("extract", pdfPath, pageIndex: 1, extractionMode: "raw");
        var json = JsonSerializer.Deserialize<JsonElement>(result);
        Assert.True(json.TryGetProperty("text", out _));
    }

    [Fact]
    public void Extract_FromSpecificPage_ShouldExtractCorrectPage()
    {
        var pdfPath = CreateMultiPagePdf("test_extract_page.pdf", 3);
        var result = _tool.Execute("extract", pdfPath, pageIndex: 2);
        var json = JsonSerializer.Deserialize<JsonElement>(result);
        Assert.Equal(2, json.GetProperty("pageIndex").GetInt32());
        Assert.Equal(3, json.GetProperty("totalPages").GetInt32());
        Assert.False(string.IsNullOrEmpty(json.GetProperty("text").GetString()));
    }

    [SkippableFact]
    public void Extract_WithUnicode_ShouldHandleUnicode()
    {
        SkipInEvaluationMode(AsposeLibraryType.Pdf, "Evaluation mode replaces text content");
        var pdfPath = CreatePdfDocument("test_unicode.pdf", "Unicode: 中文 日本語 한국어");
        var result = _tool.Execute("extract", pdfPath, pageIndex: 1);
        var json = JsonSerializer.Deserialize<JsonElement>(result);
        var text = json.GetProperty("text").GetString();
        Assert.Contains("中文", text);
        Assert.Contains("日本語", text);
    }

    [Fact]
    public void Add_ShouldAddTextToPage()
    {
        var pdfPath = CreatePdfDocument("test_add.pdf");
        var outputPath = CreateTestFilePath("test_add_output.pdf");
        var result = _tool.Execute("add", pdfPath, pageIndex: 1, outputPath: outputPath,
            text: "Added Text", x: 100, y: 700);
        Assert.True(File.Exists(outputPath));
        Assert.StartsWith("Added text to page", result);
    }

    [Fact]
    public void Add_WithFontOptions_ShouldApplyFontOptions()
    {
        var pdfPath = CreatePdfDocument("test_add_font.pdf");
        var outputPath = CreateTestFilePath("test_add_font_output.pdf");
        var result = _tool.Execute("add", pdfPath, pageIndex: 1, outputPath: outputPath,
            text: "Styled Text", fontName: "Arial", fontSize: 24, x: 100, y: 700);
        Assert.True(File.Exists(outputPath));
        Assert.StartsWith("Added text", result);
    }

    [Fact]
    public void Add_WithDefaultPosition_ShouldUseDefaults()
    {
        var pdfPath = CreatePdfDocument("test_add_default.pdf");
        var outputPath = CreateTestFilePath("test_add_default_output.pdf");
        var result = _tool.Execute("add", pdfPath, pageIndex: 1, outputPath: outputPath,
            text: "Text with default position");
        Assert.True(File.Exists(outputPath));
        Assert.StartsWith("Added text", result);
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

    [SkippableFact]
    public void Edit_WithReplaceAll_ShouldReplaceAllOccurrences()
    {
        SkipInEvaluationMode(AsposeLibraryType.Pdf, "Text editing has limitations in evaluation mode");
        var pdfPath = CreatePdfDocument("test_edit_all.pdf", "test word here. Another test word.");
        var outputPath = CreateTestFilePath("test_edit_all_output.pdf");
        var result = _tool.Execute("edit", pdfPath, pageIndex: 1, outputPath: outputPath,
            oldText: "test", newText: "replaced", replaceAll: true);
        Assert.True(File.Exists(outputPath));
        Assert.StartsWith("Replaced", result);
    }

    [SkippableFact]
    public void Edit_WithReplaceAllFalse_ShouldReplaceFirstOnly()
    {
        SkipInEvaluationMode(AsposeLibraryType.Pdf, "Text editing has limitations in evaluation mode");
        var pdfPath = CreatePdfDocument("test_edit_first.pdf", "test word. Another test word.");
        var outputPath = CreateTestFilePath("test_edit_first_output.pdf");
        var result = _tool.Execute("edit", pdfPath, pageIndex: 1, outputPath: outputPath,
            oldText: "test", newText: "replaced", replaceAll: false);
        Assert.True(File.Exists(outputPath));
        Assert.StartsWith("Replaced 1 occurrence", result);
    }

    [Theory]
    [InlineData("EXTRACT")]
    [InlineData("Extract")]
    [InlineData("extract")]
    public void Operation_ShouldBeCaseInsensitive_Extract(string operation)
    {
        var pdfPath = CreatePdfDocument($"test_case_{operation}.pdf");
        var result = _tool.Execute(operation, pdfPath, pageIndex: 1);
        Assert.Contains("pageIndex", result);
    }

    [Theory]
    [InlineData("ADD")]
    [InlineData("Add")]
    [InlineData("add")]
    public void Operation_ShouldBeCaseInsensitive_Add(string operation)
    {
        var pdfPath = CreatePdfDocument($"test_case_{operation}.pdf");
        var outputPath = CreateTestFilePath($"test_case_{operation}_output.pdf");
        var result = _tool.Execute(operation, pdfPath, pageIndex: 1, outputPath: outputPath,
            text: "Test", x: 100, y: 700);
        Assert.StartsWith("Added text", result);
    }

    [Theory]
    [InlineData("EDIT")]
    [InlineData("Edit")]
    [InlineData("edit")]
    public void Operation_ShouldBeCaseInsensitive_Edit(string operation)
    {
        var pdfPath = CreatePdfDocument($"test_case_{operation}.pdf");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute(operation, pdfPath, pageIndex: 1, oldText: "nonexistent", newText: "new"));
        Assert.Contains("not found", ex.Message);
    }

    #endregion

    #region Exception

    [Fact]
    public void Execute_WithUnknownOperation_ShouldThrowArgumentException()
    {
        var pdfPath = CreatePdfDocument("test_unknown_op.pdf");
        var ex = Assert.Throws<ArgumentException>(() => _tool.Execute("unknown", pdfPath));
        Assert.Contains("Unknown operation", ex.Message);
    }

    [Fact]
    public void Extract_WithInvalidPageIndex_ShouldThrowArgumentException()
    {
        var pdfPath = CreatePdfDocument("test_extract_invalid.pdf");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("extract", pdfPath, pageIndex: 99));
        Assert.Contains("pageIndex must be between", ex.Message);
    }

    [Fact]
    public void Add_WithInvalidPageIndex_ShouldThrowArgumentException()
    {
        var pdfPath = CreatePdfDocument("test_add_invalid.pdf");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("add", pdfPath, pageIndex: 99, text: "Test"));
        Assert.Contains("pageIndex must be between", ex.Message);
    }

    [Fact]
    public void Add_WithMissingText_ShouldThrowArgumentException()
    {
        var pdfPath = CreatePdfDocument("test_add_notext.pdf");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("add", pdfPath, pageIndex: 1, text: null));
        Assert.Contains("text is required", ex.Message);
    }

    [Fact]
    public void Add_WithEmptyText_ShouldThrowArgumentException()
    {
        var pdfPath = CreatePdfDocument("test_add_empty.pdf");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("add", pdfPath, pageIndex: 1, text: ""));
        Assert.Contains("text is required", ex.Message);
    }

    [Fact]
    public void Edit_WithInvalidPageIndex_ShouldThrowArgumentException()
    {
        var pdfPath = CreatePdfDocument("test_edit_invalid.pdf");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("edit", pdfPath, pageIndex: 99, oldText: "old", newText: "new"));
        Assert.Contains("pageIndex must be between", ex.Message);
    }

    [Fact]
    public void Edit_WithMissingOldText_ShouldThrowArgumentException()
    {
        var pdfPath = CreatePdfDocument("test_edit_noold.pdf");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("edit", pdfPath, pageIndex: 1, oldText: null, newText: "new"));
        Assert.Contains("oldText is required", ex.Message);
    }

    [Fact]
    public void Edit_WithMissingNewText_ShouldThrowArgumentException()
    {
        var pdfPath = CreatePdfDocument("test_edit_nonew.pdf");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("edit", pdfPath, pageIndex: 1, oldText: "old", newText: null));
        Assert.Contains("newText is required", ex.Message);
    }

    [Fact]
    public void Edit_WithTextNotFound_ShouldThrowArgumentException()
    {
        var pdfPath = CreatePdfDocument("test_edit_notfound.pdf");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("edit", pdfPath, pageIndex: 1, oldText: "nonexistent_12345", newText: "new"));
        Assert.Contains("not found", ex.Message);
    }

    [Fact]
    public void Execute_WithNoPathOrSessionId_ShouldThrowException()
    {
        Assert.ThrowsAny<Exception>(() => _tool.Execute("extract"));
    }

    #endregion

    #region Session

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
    public void Extract_WithSessionId_AndIncludeFontInfo_ShouldReturnFragments()
    {
        var pdfPath = CreatePdfDocument("test_session_font.pdf");
        var sessionId = OpenSession(pdfPath);
        var result = _tool.Execute("extract", sessionId: sessionId, pageIndex: 1, includeFontInfo: true);
        var json = JsonSerializer.Deserialize<JsonElement>(result);
        Assert.True(json.TryGetProperty("fragments", out _));
    }

    [Fact]
    public void Add_WithSessionId_ShouldAddToSession()
    {
        var pdfPath = CreatePdfDocument("test_session_add.pdf");
        var sessionId = OpenSession(pdfPath);
        var result = _tool.Execute("add", sessionId: sessionId, pageIndex: 1,
            text: "Session Text", x: 100, y: 700);
        Assert.StartsWith("Added text", result);
        Assert.Contains("session", result);
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
        Assert.Contains("session", result);
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
        Assert.DoesNotContain("Path Content", text);
    }

    #endregion
}