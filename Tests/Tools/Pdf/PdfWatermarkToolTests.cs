using Aspose.Pdf;
using Aspose.Pdf.Text;
using AsposeMcpServer.Tests.Helpers;
using AsposeMcpServer.Tools.Pdf;

namespace AsposeMcpServer.Tests.Tools.Pdf;

public class PdfWatermarkToolTests : PdfTestBase
{
    private readonly PdfWatermarkTool _tool;

    public PdfWatermarkToolTests()
    {
        _tool = new PdfWatermarkTool(SessionManager);
    }

    private string CreatePdfDocument(string fileName, int pageCount = 1)
    {
        var filePath = CreateTestFilePath(fileName);
        var document = new Document();
        for (var i = 0; i < pageCount; i++)
        {
            var page = document.Pages.Add();
            page.Paragraphs.Add(new TextFragment($"Page {i + 1} Content"));
        }

        document.Save(filePath);
        return filePath;
    }

    #region General Tests

    [Fact]
    public void AddWatermark_WithText_ShouldAddWatermark()
    {
        var pdfPath = CreatePdfDocument("test_add_watermark.pdf");
        var outputPath = CreateTestFilePath("test_add_watermark_output.pdf");
        var result = _tool.Execute(
            "add",
            text: "Confidential",
            path: pdfPath,
            outputPath: outputPath);
        Assert.True(File.Exists(outputPath), "PDF file should be created");
        Assert.Contains("Watermark added to 1 page(s)", result);
    }

    [Fact]
    public void AddWatermark_WithFontOptions_ShouldApplyFontOptions()
    {
        var pdfPath = CreatePdfDocument("test_add_watermark_font.pdf");
        var outputPath = CreateTestFilePath("test_add_watermark_font_output.pdf");
        var result = _tool.Execute(
            "add",
            text: "Watermark",
            path: pdfPath,
            outputPath: outputPath,
            fontName: "Arial",
            fontSize: 72);
        Assert.True(File.Exists(outputPath), "PDF file should be created");
        Assert.Contains("Watermark added", result);
    }

    [Fact]
    public void AddWatermark_WithOpacity_ShouldApplyOpacity()
    {
        var pdfPath = CreatePdfDocument("test_add_watermark_opacity.pdf");
        var outputPath = CreateTestFilePath("test_add_watermark_opacity_output.pdf");
        _tool.Execute(
            "add",
            text: "Watermark",
            path: pdfPath,
            outputPath: outputPath,
            opacity: 0.5);
        Assert.True(File.Exists(outputPath), "PDF file should be created");
    }

    [Fact]
    public void AddWatermark_WithRotation_ShouldApplyRotation()
    {
        var pdfPath = CreatePdfDocument("test_add_watermark_rotation.pdf");
        var outputPath = CreateTestFilePath("test_add_watermark_rotation_output.pdf");
        _tool.Execute(
            "add",
            text: "Watermark",
            path: pdfPath,
            outputPath: outputPath,
            rotation: 45);
        Assert.True(File.Exists(outputPath), "PDF file should be created");
    }

    [Fact]
    public void AddWatermark_WithAlignment_ShouldApplyAlignment()
    {
        var pdfPath = CreatePdfDocument("test_add_watermark_alignment.pdf");
        var outputPath = CreateTestFilePath("test_add_watermark_alignment_output.pdf");
        _tool.Execute(
            "add",
            text: "Watermark",
            path: pdfPath,
            outputPath: outputPath,
            horizontalAlignment: "Left",
            verticalAlignment: "Top");
        Assert.True(File.Exists(outputPath), "PDF file should be created");
    }

    [Fact]
    public void AddWatermark_WithColor_ShouldApplyColor()
    {
        var pdfPath = CreatePdfDocument("test_add_watermark_color.pdf");
        var outputPath = CreateTestFilePath("test_add_watermark_color_output.pdf");
        var result = _tool.Execute(
            "add",
            text: "URGENT",
            path: pdfPath,
            outputPath: outputPath,
            color: "Red");
        Assert.True(File.Exists(outputPath), "PDF file should be created");
        Assert.Contains("Watermark added", result);
    }

    [Fact]
    public void AddWatermark_WithHexColor_ShouldApplyHexColor()
    {
        var pdfPath = CreatePdfDocument("test_add_watermark_hex.pdf");
        var outputPath = CreateTestFilePath("test_add_watermark_hex_output.pdf");
        _tool.Execute(
            "add",
            text: "Custom Color",
            path: pdfPath,
            outputPath: outputPath,
            color: "#FF5500");
        Assert.True(File.Exists(outputPath), "PDF file should be created");
    }

    [SkippableFact]
    public void AddWatermark_WithPageRange_ShouldApplyToSpecificPages()
    {
        // Skip in evaluation mode - 5 pages exceeds 4-page limit
        SkipInEvaluationMode(AsposeLibraryType.Pdf, "5 pages exceeds 4-page limit in evaluation mode");
        var pdfPath = CreatePdfDocument("test_add_watermark_range.pdf", 5);
        var outputPath = CreateTestFilePath("test_add_watermark_range_output.pdf");
        var result = _tool.Execute(
            "add",
            text: "Selected Pages",
            path: pdfPath,
            outputPath: outputPath,
            pageRange: "1,3,5");
        Assert.True(File.Exists(outputPath), "PDF file should be created");
        Assert.Contains("Watermark added to 3 page(s)", result);
    }

    [SkippableFact]
    public void AddWatermark_WithPageRangeHyphen_ShouldApplyToRange()
    {
        // Skip in evaluation mode - 10 pages exceeds 4-page limit
        SkipInEvaluationMode(AsposeLibraryType.Pdf, "10 pages exceeds 4-page limit in evaluation mode");
        var pdfPath = CreatePdfDocument("test_add_watermark_hyphen.pdf", 10);
        var outputPath = CreateTestFilePath("test_add_watermark_hyphen_output.pdf");
        var result = _tool.Execute(
            "add",
            text: "Range Pages",
            path: pdfPath,
            outputPath: outputPath,
            pageRange: "2-5");
        Assert.True(File.Exists(outputPath), "PDF file should be created");
        Assert.Contains("Watermark added to 4 page(s)", result);
    }

    [SkippableFact]
    public void AddWatermark_WithMixedPageRange_ShouldApplyCorrectly()
    {
        // Skip in evaluation mode - 10 pages exceeds 4-page limit
        SkipInEvaluationMode(AsposeLibraryType.Pdf, "10 pages exceeds 4-page limit in evaluation mode");
        var pdfPath = CreatePdfDocument("test_add_watermark_mixed.pdf", 10);
        var outputPath = CreateTestFilePath("test_add_watermark_mixed_output.pdf");
        var result = _tool.Execute(
            "add",
            text: "Mixed Range",
            path: pdfPath,
            outputPath: outputPath,
            pageRange: "1,3-5,8");
        Assert.True(File.Exists(outputPath), "PDF file should be created");
        Assert.Contains("Watermark added to 5 page(s)", result);
    }

    [Fact]
    public void AddWatermark_WithIsBackground_ShouldSetBackground()
    {
        var pdfPath = CreatePdfDocument("test_add_watermark_bg.pdf");
        var outputPath = CreateTestFilePath("test_add_watermark_bg_output.pdf");
        _tool.Execute(
            "add",
            text: "Background Watermark",
            path: pdfPath,
            outputPath: outputPath,
            isBackground: true);
        Assert.True(File.Exists(outputPath), "PDF file should be created");
    }

    [Fact]
    public void AddWatermark_WithAllOptions_ShouldApplyAllOptions()
    {
        var pdfPath = CreatePdfDocument("test_add_watermark_all.pdf", 3);
        var outputPath = CreateTestFilePath("test_add_watermark_all_output.pdf");
        var result = _tool.Execute(
            "add",
            text: "Confidential",
            path: pdfPath,
            outputPath: outputPath,
            fontName: "Arial",
            fontSize: 72,
            opacity: 0.3,
            rotation: 45,
            color: "Red",
            pageRange: "1-2",
            isBackground: true,
            horizontalAlignment: "Center",
            verticalAlignment: "Center");
        Assert.True(File.Exists(outputPath), "PDF file should be created");
        Assert.Contains("Watermark added to 2 page(s)", result);
    }

    [Fact]
    public void AddWatermark_WithInvalidPageRange_ShouldThrowArgumentException()
    {
        var pdfPath = CreatePdfDocument("test_invalid_range.pdf", 3);
        var exception = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("add", text: "Test", path: pdfPath, pageRange: "invalid"));
        Assert.Contains("Invalid page number", exception.Message);
    }

    [Fact]
    public void AddWatermark_WithOutOfBoundsPageRange_ShouldThrowArgumentException()
    {
        var pdfPath = CreatePdfDocument("test_oob_range.pdf", 3);
        var exception = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("add", text: "Test", path: pdfPath, pageRange: "1,5"));
        Assert.Contains("out of bounds", exception.Message);
    }

    [SkippableFact]
    public void AddWatermark_WithInvalidRangeFormat_ShouldThrowArgumentException()
    {
        // Skip in evaluation mode - 5 pages exceeds 4-page limit
        SkipInEvaluationMode(AsposeLibraryType.Pdf, "5 pages exceeds 4-page limit in evaluation mode");
        var pdfPath = CreatePdfDocument("test_invalid_format.pdf", 5);
        var exception = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("add", text: "Test", path: pdfPath, pageRange: "3-1"));
        Assert.Contains("out of bounds", exception.Message);
    }

    [Fact]
    public void Execute_WithUnknownOperation_ShouldThrowArgumentException()
    {
        var pdfPath = CreatePdfDocument("test_unknown_op.pdf");
        var exception = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("unknown", text: "Test", path: pdfPath));
        Assert.Contains("Unknown operation", exception.Message);
    }

    [SkippableFact]
    public void AddWatermark_WithMultiplePages_ShouldApplyToAllPages()
    {
        // Skip in evaluation mode - 5 pages exceeds 4-page limit
        SkipInEvaluationMode(AsposeLibraryType.Pdf, "5 pages exceeds 4-page limit in evaluation mode");
        var pdfPath = CreatePdfDocument("test_multi_page.pdf", 5);
        var outputPath = CreateTestFilePath("test_multi_page_output.pdf");
        var result = _tool.Execute(
            "add",
            text: "All Pages",
            path: pdfPath,
            outputPath: outputPath);
        Assert.True(File.Exists(outputPath), "PDF file should be created");
        Assert.Contains("Watermark added to 5 page(s)", result);
    }

    [Fact]
    public void AddWatermark_WithUnknownColor_ShouldUseGray()
    {
        var pdfPath = CreatePdfDocument("test_unknown_color.pdf");
        var outputPath = CreateTestFilePath("test_unknown_color_output.pdf");
        _tool.Execute(
            "add",
            text: "Unknown Color",
            path: pdfPath,
            outputPath: outputPath,
            color: "InvalidColor");
        Assert.True(File.Exists(outputPath), "PDF file should be created");
    }

    #endregion

    #region Exception Tests

    [Fact]
    public void Execute_WithNoPathOrSessionId_ShouldThrowException()
    {
        Assert.ThrowsAny<Exception>(() => _tool.Execute(
            "add",
            text: "Test",
            path: null,
            sessionId: null));
    }

    [Fact]
    public void AddWatermark_WithMissingText_ShouldThrowArgumentException()
    {
        var pdfPath = CreatePdfDocument("test_missing_text.pdf");
        var exception = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("add", pdfPath, text: null));
        Assert.Contains("text is required", exception.Message);
    }

    #endregion

    #region Session ID Tests

    [Fact]
    public void AddWatermark_WithSessionId_ShouldModifyInMemory()
    {
        var pdfPath = CreatePdfDocument("test_session_add_watermark.pdf");
        var sessionId = OpenSession(pdfPath);
        var result = _tool.Execute(
            "add",
            text: "Confidential",
            sessionId: sessionId);
        Assert.Contains("Watermark added to 1 page(s)", result);
        var document = SessionManager.GetDocument<Document>(sessionId);
        Assert.NotNull(document);
        Assert.Single(document.Pages);
    }

    [Fact]
    public void AddWatermark_WithSessionId_AndOptions_ShouldApplyOptionsInMemory()
    {
        var pdfPath = CreatePdfDocument("test_session_watermark_options.pdf", 2);
        var sessionId = OpenSession(pdfPath);
        var result = _tool.Execute(
            "add",
            text: "DRAFT",
            sessionId: sessionId,
            fontName: "Arial",
            fontSize: 72,
            opacity: 0.5,
            rotation: 45,
            color: "Red");
        Assert.Contains("Watermark added to 2 page(s)", result);
        var document = SessionManager.GetDocument<Document>(sessionId);
        Assert.NotNull(document);
        Assert.Equal(2, document.Pages.Count);
    }

    [Fact]
    public void AddWatermark_WithSessionId_AndPageRange_ShouldApplyToSpecificPages()
    {
        var pdfPath = CreatePdfDocument("test_session_watermark_range.pdf", 3);
        var sessionId = OpenSession(pdfPath);
        var result = _tool.Execute(
            "add",
            text: "Selected",
            sessionId: sessionId,
            pageRange: "1,3");
        Assert.Contains("Watermark added to 2 page(s)", result);
        var document = SessionManager.GetDocument<Document>(sessionId);
        Assert.NotNull(document);
        Assert.Equal(3, document.Pages.Count);
    }

    #endregion
}