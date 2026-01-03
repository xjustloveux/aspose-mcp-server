using Aspose.Words;
using AsposeMcpServer.Tests.Helpers;
using AsposeMcpServer.Tools.Word;

namespace AsposeMcpServer.Tests.Tools.Word;

public class WordHeaderFooterToolTests : WordTestBase
{
    private readonly WordHeaderFooterTool _tool;

    public WordHeaderFooterToolTests()
    {
        _tool = new WordHeaderFooterTool(SessionManager);
    }

    #region General Tests

    [Fact]
    public void SetHeaderText_ShouldSetHeaderText()
    {
        var docPath = CreateWordDocument("test_set_header_text.docx");
        var outputPath = CreateTestFilePath("test_set_header_text_output.docx");
        _tool.Execute("set_header_text", docPath, outputPath: outputPath,
            headerLeft: "Left Header", headerCenter: "Center Header", headerRight: "Right Header");
        var doc = new Document(outputPath);
        var header = doc.FirstSection.HeadersFooters[HeaderFooterType.HeaderPrimary];
        Assert.NotNull(header);
        Assert.Contains("Left", header.GetText(), StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void SetFooterText_ShouldSetFooterText()
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
    public void GetHeadersFooters_ShouldReturnHeadersFooters()
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
        Assert.NotNull(result);
        Assert.NotEmpty(result);
        Assert.Contains("Header", result, StringComparison.OrdinalIgnoreCase);
    }

    [SkippableFact]
    public void SetHeaderImage_ShouldSetHeaderImage()
    {
        // Skip test if running in evaluation mode as image operations may be limited
        SkipInEvaluationMode(AsposeLibraryType.Words, "Image operations may be limited in evaluation mode");
        var docPath = CreateWordDocument("test_set_header_image.docx");

        // Create a simple test image
        var imagePath = CreateTestFilePath("test_header_image.png");
        CreateTestImage(imagePath);

        var outputPath = CreateTestFilePath("test_set_header_image_output.docx");
        _tool.Execute("set_header_image", docPath, outputPath: outputPath,
            imagePath: imagePath, alignment: "center");
        Assert.True(File.Exists(outputPath), "Output document should be created");
    }

    [Fact]
    public void SetHeaderLine_ShouldSetHeaderLine()
    {
        var docPath = CreateWordDocument("test_set_header_line.docx");
        var outputPath = CreateTestFilePath("test_set_header_line_output.docx");
        _tool.Execute("set_header_line", docPath, outputPath: outputPath, lineStyle: "single");
        Assert.True(File.Exists(outputPath), "Output document should be created");
        var doc = new Document(outputPath);
        var header = doc.FirstSection.HeadersFooters[HeaderFooterType.HeaderPrimary];
        Assert.NotNull(header);
    }

    [Fact]
    public void SetFooterLine_ShouldSetFooterLine()
    {
        var docPath = CreateWordDocument("test_set_footer_line.docx");
        var outputPath = CreateTestFilePath("test_set_footer_line_output.docx");
        _tool.Execute("set_footer_line", docPath, outputPath: outputPath, lineStyle: "single");
        Assert.True(File.Exists(outputPath), "Output document should be created");
        var doc = new Document(outputPath);
        var footer = doc.FirstSection.HeadersFooters[HeaderFooterType.FooterPrimary];
        Assert.NotNull(footer);
    }

    [Fact]
    public void SetHeaderFooter_ShouldSetBoth()
    {
        var docPath = CreateWordDocument("test_set_header_footer.docx");
        var outputPath = CreateTestFilePath("test_set_header_footer_output.docx");
        _tool.Execute("set_header_footer", docPath, outputPath: outputPath,
            headerLeft: "Left Header", footerCenter: "Center Footer");
        Assert.True(File.Exists(outputPath), "Output document should be created");
        var doc = new Document(outputPath);
        var header = doc.FirstSection.HeadersFooters[HeaderFooterType.HeaderPrimary];
        var footer = doc.FirstSection.HeadersFooters[HeaderFooterType.FooterPrimary];
        Assert.NotNull(header);
        Assert.NotNull(footer);
        Assert.Contains("Left", header.GetText(), StringComparison.OrdinalIgnoreCase);
        Assert.Contains("Center", footer.GetText(), StringComparison.OrdinalIgnoreCase);
    }

    /// <summary>
    ///     Creates a simple test image file for testing
    /// </summary>
    private void CreateTestImage(string imagePath)
    {
        // Create a minimal valid PNG file (1x1 pixel white)
        byte[] pngBytes =
        [
            0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A, // PNG signature
            0x00, 0x00, 0x00, 0x0D, 0x49, 0x48, 0x44, 0x52, // IHDR chunk
            0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, // 1x1 dimensions
            0x08, 0x02, 0x00, 0x00, 0x00, 0x90, 0x77, 0x53, // bit depth, color type, etc.
            0xDE, 0x00, 0x00, 0x00, 0x0C, 0x49, 0x44, 0x41, // IDAT chunk
            0x54, 0x08, 0xD7, 0x63, 0xF8, 0xFF, 0xFF, 0xFF, // compressed data
            0x00, 0x05, 0xFE, 0x02, 0xFE, 0xDC, 0xCC, 0x59, //
            0xE7, 0x00, 0x00, 0x00, 0x00, 0x49, 0x45, 0x4E, // IEND chunk
            0x44, 0xAE, 0x42, 0x60, 0x82
        ];
        File.WriteAllBytes(imagePath, pngBytes);
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
        Assert.Contains("unknown_operation", ex.Message);
    }

    [Fact]
    public void SetHeaderImage_WithMissingImagePath_ShouldThrowArgumentException()
    {
        var docPath = CreateWordDocument("test_header_missing_image.docx");
        var outputPath = CreateTestFilePath("test_header_missing_image_output.docx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("set_header_image", docPath, outputPath: outputPath, imagePath: null));

        Assert.Contains("imagePath cannot be null or empty", ex.Message);
    }

    [Fact]
    public void SetHeaderImage_WithNonExistentImagePath_ShouldThrowFileNotFoundException()
    {
        var docPath = CreateWordDocument("test_header_nonexistent_image.docx");
        var outputPath = CreateTestFilePath("test_header_nonexistent_image_output.docx");
        Assert.Throws<FileNotFoundException>(() =>
            _tool.Execute("set_header_image", docPath, outputPath: outputPath,
                imagePath: "C:\\nonexistent\\image.png"));
    }

    #endregion

    #region Session ID Tests

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
        Assert.Contains("Header", result, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void SetHeaderText_WithSessionId_ShouldSetHeaderInMemory()
    {
        var docPath = CreateWordDocument("test_session_set_header.docx");
        var sessionId = OpenSession(docPath);
        var result = _tool.Execute("set_header_text", sessionId: sessionId,
            headerLeft: "Session Left", headerCenter: "Session Center");
        Assert.Contains("Header", result, StringComparison.OrdinalIgnoreCase);

        // Verify in-memory document has the header
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
        Assert.Contains("Footer", result, StringComparison.OrdinalIgnoreCase);

        // Verify in-memory document has the footer
        var doc = SessionManager.GetDocument<Document>(sessionId);
        var footer = doc.FirstSection.HeadersFooters[HeaderFooterType.FooterPrimary];
        Assert.NotNull(footer);
        Assert.Contains("Session Footer", footer.GetText(), StringComparison.OrdinalIgnoreCase);
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

        // Act - provide both path and sessionId
        var result = _tool.Execute("get", docPath1, sessionId);

        // Assert - should use sessionId, returning Session Header not Path Header
        Assert.Contains("Session Header Unique", result);
        Assert.DoesNotContain("Path Header", result);
    }

    #endregion
}