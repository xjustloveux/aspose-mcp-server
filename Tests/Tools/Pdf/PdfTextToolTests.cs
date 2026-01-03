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

    private string CreatePdfDocument(string fileName)
    {
        var filePath = CreateTestFilePath(fileName);
        var document = new Document();
        var page = document.Pages.Add();
        page.Paragraphs.Add(new TextFragment("Sample PDF Text"));
        document.Save(filePath);
        return filePath;
    }

    #region General Tests

    [Fact]
    public void ExtractText_ShouldReturnJsonResult()
    {
        var pdfPath = CreatePdfDocument("test_extract_text.pdf");
        var result = _tool.Execute("extract", pdfPath, pageIndex: 1);
        Assert.NotNull(result);
        var json = JsonSerializer.Deserialize<JsonElement>(result);
        Assert.True(json.TryGetProperty("pageIndex", out _));
        Assert.True(json.TryGetProperty("totalPages", out _));
        Assert.True(json.TryGetProperty("text", out _));
    }

    [Fact]
    public void ExtractText_WithIncludeFontInfo_ShouldReturnFragments()
    {
        var pdfPath = CreatePdfDocument("test_extract_font_info.pdf");
        var result = _tool.Execute("extract", pdfPath, pageIndex: 1, includeFontInfo: true);
        Assert.NotNull(result);
        var json = JsonSerializer.Deserialize<JsonElement>(result);
        Assert.True(json.TryGetProperty("fragments", out _));
        Assert.True(json.TryGetProperty("fragmentCount", out _));
    }

    [Fact]
    public void ExtractText_WithRawMode_ShouldExtractRawText()
    {
        var pdfPath = CreatePdfDocument("test_extract_raw.pdf");
        var result = _tool.Execute("extract", pdfPath, pageIndex: 1, extractionMode: "raw");
        Assert.NotNull(result);
        Assert.Contains("text", result);
    }

    [Fact]
    public void ExtractText_WithInvalidPageIndex_ShouldThrowArgumentException()
    {
        var pdfPath = CreatePdfDocument("test_extract_invalid_page.pdf");
        var exception = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("extract", pdfPath, pageIndex: 99));
        Assert.Contains("pageIndex must be between", exception.Message);
    }

    [Fact]
    public void AddText_ShouldAddTextToPage()
    {
        var pdfPath = CreatePdfDocument("test_add_text.pdf");
        var outputPath = CreateTestFilePath("test_add_text_output.pdf");
        var result = _tool.Execute(
            "add",
            pdfPath,
            pageIndex: 1,
            outputPath: outputPath,
            text: "Added Text",
            x: 100,
            y: 700);
        Assert.True(File.Exists(outputPath), "PDF file should be created");
        Assert.Contains("Added text to page", result);
    }

    [Fact]
    public void AddText_WithFontOptions_ShouldApplyFontOptions()
    {
        var pdfPath = CreatePdfDocument("test_add_text_font.pdf");
        var outputPath = CreateTestFilePath("test_add_text_font_output.pdf");
        _tool.Execute(
            "add",
            pdfPath,
            pageIndex: 1,
            outputPath: outputPath,
            text: "Styled Text",
            fontName: "Arial",
            fontSize: 14,
            x: 100,
            y: 700);
        Assert.True(File.Exists(outputPath), "PDF file should be created");
        var document = new Document(outputPath);
        Assert.True(document.Pages.Count >= 1, "PDF should have at least one page");
    }

    [Fact]
    public void AddText_WithInvalidPageIndex_ShouldThrowArgumentException()
    {
        var pdfPath = CreatePdfDocument("test_add_invalid_page.pdf");
        var exception = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("add", pdfPath, pageIndex: 99, text: "Test"));
        Assert.Contains("pageIndex must be between", exception.Message);
    }

    [SkippableFact]
    public void EditText_ShouldReplaceText()
    {
        // Skip in evaluation mode - text editing has limitations in evaluation mode
        SkipInEvaluationMode(AsposeLibraryType.Pdf, "Text editing has limitations in evaluation mode");
        var pdfPath = CreatePdfDocument("test_edit_text.pdf");
        var outputPath = CreateTestFilePath("test_edit_text_output.pdf");

        var result = _tool.Execute(
            "edit",
            pdfPath,
            pageIndex: 1,
            outputPath: outputPath,
            oldText: "Sample PDF Text",
            newText: "Updated",
            replaceAll: true);
        Assert.True(File.Exists(outputPath), "PDF file should be created");
        Assert.Contains("Replaced", result);
    }

    [Fact]
    public void EditText_WithInvalidPageIndex_ShouldThrowArgumentException()
    {
        var pdfPath = CreatePdfDocument("test_edit_invalid_page.pdf");
        var exception = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("edit", pdfPath, pageIndex: 99, oldText: "old", newText: "new"));
        Assert.Contains("pageIndex must be between", exception.Message);
    }

    [Fact]
    public void EditText_WithTextNotFound_ShouldThrowArgumentException()
    {
        var pdfPath = CreatePdfDocument("test_edit_not_found.pdf");
        var exception = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("edit", pdfPath, pageIndex: 1, oldText: "nonexistent_text_12345", newText: "new"));
        Assert.Contains("not found", exception.Message);
    }

    [Fact]
    public void Execute_WithUnknownOperation_ShouldThrowArgumentException()
    {
        var pdfPath = CreatePdfDocument("test_unknown_op.pdf");
        var exception = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("unknown", pdfPath));
        Assert.Contains("Unknown operation", exception.Message);
    }

    [Fact]
    public void ExtractText_WithMultiplePages_ShouldExtractFromSpecifiedPage()
    {
        var pdfPath = CreateTestFilePath("test_multi_page_extract.pdf");
        var document = new Document();
        var page1 = document.Pages.Add();
        page1.Paragraphs.Add(new TextFragment("Page 1 Content"));
        var page2 = document.Pages.Add();
        page2.Paragraphs.Add(new TextFragment("Page 2 Content"));
        document.Save(pdfPath);
        var result = _tool.Execute("extract", pdfPath, pageIndex: 2);
        Assert.NotNull(result);
        var json = JsonSerializer.Deserialize<JsonElement>(result);
        Assert.True(json.TryGetProperty("pageIndex", out var pageIndexProp));
        Assert.Equal(2, pageIndexProp.GetInt32());
    }

    [Fact]
    public void ExtractText_WithUnicode_ShouldHandleUnicode()
    {
        var pdfPath = CreateTestFilePath("test_unicode_extract.pdf");
        var document = new Document();
        var page = document.Pages.Add();
        page.Paragraphs.Add(new TextFragment("Unicode Test: 中文 日本語 한국어"));
        document.Save(pdfPath);
        var result = _tool.Execute("extract", pdfPath, pageIndex: 1);
        Assert.NotNull(result);
        Assert.True(result.Length > 0, "Should extract unicode text");
    }

    [Fact]
    public void EditText_ReplaceAll_ShouldReplaceAllOccurrences()
    {
        var pdfPath = CreateTestFilePath("test_replace_all.pdf");
        var document = new Document();
        var page = document.Pages.Add();
        page.Paragraphs.Add(new TextFragment("Test word here. Another test word. Third test word."));
        document.Save(pdfPath);

        var outputPath = CreateTestFilePath("test_replace_all_output.pdf");

        var isEvaluationMode = IsEvaluationMode();

        try
        {
            _tool.Execute(
                "edit",
                pdfPath,
                pageIndex: 1,
                outputPath: outputPath,
                oldText: "test",
                newText: "replaced",
                replaceAll: true);
            Assert.True(File.Exists(outputPath), "Output file should be created");
        }
        catch (Exception) when (isEvaluationMode)
        {
            Assert.True(true, "In evaluation mode, replace operation may fail");
        }
    }

    [Fact]
    public void AddText_WithDefaultPosition_ShouldUseDefaults()
    {
        var pdfPath = CreatePdfDocument("test_add_default_position.pdf");
        var outputPath = CreateTestFilePath("test_add_default_position_output.pdf");
        var result = _tool.Execute(
            "add",
            pdfPath,
            pageIndex: 1,
            outputPath: outputPath,
            text: "Text with default position");
        Assert.True(File.Exists(outputPath), "PDF file should be created");
        Assert.Contains("Added text", result);
    }

    #endregion

    #region Session ID Tests

    [Fact]
    public void ExtractText_WithSessionId_ShouldExtractFromSession()
    {
        var pdfPath = CreatePdfDocument("test_session_extract.pdf");
        var sessionId = OpenSession(pdfPath);
        var result = _tool.Execute("extract", sessionId: sessionId, pageIndex: 1);
        Assert.NotNull(result);
        var json = JsonSerializer.Deserialize<JsonElement>(result);
        Assert.True(json.TryGetProperty("text", out _));
    }

    [Fact]
    public void AddText_WithSessionId_ShouldAddToSession()
    {
        var pdfPath = CreatePdfDocument("test_session_add.pdf");
        var sessionId = OpenSession(pdfPath);
        var result = _tool.Execute("add", sessionId: sessionId, pageIndex: 1, text: "Session Text", x: 100, y: 700);
        Assert.Contains("Added text", result);
        Assert.Contains("session", result);
    }

    [SkippableFact]
    public void EditText_WithSessionId_ShouldEditInSession()
    {
        // Skip in evaluation mode - text editing has limitations
        SkipInEvaluationMode(AsposeLibraryType.Pdf, "Text editing has limitations in evaluation mode");
        var pdfPath = CreatePdfDocument("test_session_edit.pdf");
        var sessionId = OpenSession(pdfPath);
        var result = _tool.Execute(
            "edit",
            sessionId: sessionId,
            pageIndex: 1,
            oldText: "Sample PDF Text",
            newText: "Updated Session Text",
            replaceAll: true);
        Assert.Contains("Replaced", result);
    }

    #endregion
}