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
        using var document = new Document();
        for (var i = 0; i < pageCount; i++)
        {
            var page = document.Pages.Add();
            page.Paragraphs.Add(new TextFragment($"Page {i + 1} Content"));
        }

        document.Save(filePath);
        return filePath;
    }

    #region General

    [Fact]
    public void Add_WithText_ShouldAddWatermark()
    {
        const string watermarkText = "Confidential";
        var pdfPath = CreatePdfDocument("test_add.pdf");
        var outputPath = CreateTestFilePath("test_add_output.pdf");

        var result = _tool.Execute("add", text: watermarkText, path: pdfPath, outputPath: outputPath);

        Assert.True(File.Exists(outputPath));
        Assert.StartsWith("Watermark added to 1 page(s)", result);

        using var outputDoc = new Document(outputPath);
        var textAbsorber = new TextAbsorber();
        outputDoc.Pages.Accept(textAbsorber);
        Assert.Contains(watermarkText, textAbsorber.Text);
    }

    [Fact]
    public void Add_WithFontOptions_ShouldApplyFontOptions()
    {
        const string watermarkText = "Watermark";
        var pdfPath = CreatePdfDocument("test_font.pdf");
        var outputPath = CreateTestFilePath("test_font_output.pdf");

        var result = _tool.Execute("add", text: watermarkText, path: pdfPath, outputPath: outputPath, fontName: "Arial",
            fontSize: 72);

        Assert.True(File.Exists(outputPath));
        Assert.StartsWith("Watermark added to 1 page(s)", result);

        using var outputDoc = new Document(outputPath);
        var textAbsorber = new TextAbsorber();
        outputDoc.Pages.Accept(textAbsorber);
        Assert.Contains(watermarkText, textAbsorber.Text);
    }

    [Fact]
    public void Add_WithOpacity_ShouldApplyOpacity()
    {
        const string watermarkText = "Watermark";
        var pdfPath = CreatePdfDocument("test_opacity.pdf");
        var outputPath = CreateTestFilePath("test_opacity_output.pdf");

        _tool.Execute("add", text: watermarkText, path: pdfPath, outputPath: outputPath, opacity: 0.5);

        Assert.True(File.Exists(outputPath));

        using var outputDoc = new Document(outputPath);
        var textAbsorber = new TextAbsorber();
        outputDoc.Pages.Accept(textAbsorber);
        Assert.Contains(watermarkText, textAbsorber.Text);
    }

    [Fact]
    public void Add_WithRotation_ShouldApplyRotation()
    {
        const string watermarkText = "Watermark";
        var pdfPath = CreatePdfDocument("test_rotation.pdf");
        var outputPath = CreateTestFilePath("test_rotation_output.pdf");

        _tool.Execute("add", text: watermarkText, path: pdfPath, outputPath: outputPath, rotation: 45);

        Assert.True(File.Exists(outputPath));

        using var outputDoc = new Document(outputPath);
        var textAbsorber = new TextAbsorber();
        outputDoc.Pages.Accept(textAbsorber);
        Assert.Contains(watermarkText, textAbsorber.Text);
    }

    [Fact]
    public void Add_WithAlignment_ShouldApplyAlignment()
    {
        const string watermarkText = "Watermark";
        var pdfPath = CreatePdfDocument("test_alignment.pdf");
        var outputPath = CreateTestFilePath("test_alignment_output.pdf");

        _tool.Execute("add", text: watermarkText, path: pdfPath, outputPath: outputPath, horizontalAlignment: "Left",
            verticalAlignment: "Top");

        Assert.True(File.Exists(outputPath));

        using var outputDoc = new Document(outputPath);
        var textAbsorber = new TextAbsorber();
        outputDoc.Pages.Accept(textAbsorber);
        Assert.Contains(watermarkText, textAbsorber.Text);
    }

    [Fact]
    public void Add_WithNamedColor_ShouldApplyColor()
    {
        const string watermarkText = "URGENT";
        var pdfPath = CreatePdfDocument("test_color.pdf");
        var outputPath = CreateTestFilePath("test_color_output.pdf");

        var result = _tool.Execute("add", text: watermarkText, path: pdfPath, outputPath: outputPath, color: "Red");

        Assert.True(File.Exists(outputPath));
        Assert.StartsWith("Watermark added to 1 page(s)", result);

        using var outputDoc = new Document(outputPath);
        var textAbsorber = new TextAbsorber();
        outputDoc.Pages.Accept(textAbsorber);
        Assert.Contains(watermarkText, textAbsorber.Text);
    }

    [Fact]
    public void Add_WithHexColor_ShouldApplyHexColor()
    {
        const string watermarkText = "Custom Color";
        var pdfPath = CreatePdfDocument("test_hex.pdf");
        var outputPath = CreateTestFilePath("test_hex_output.pdf");

        _tool.Execute("add", text: watermarkText, path: pdfPath, outputPath: outputPath, color: "#FF5500");

        Assert.True(File.Exists(outputPath));

        using var outputDoc = new Document(outputPath);
        var textAbsorber = new TextAbsorber();
        outputDoc.Pages.Accept(textAbsorber);
        Assert.Contains(watermarkText, textAbsorber.Text);
    }

    [Fact]
    public void Add_WithUnknownColor_ShouldUseGray()
    {
        const string watermarkText = "Unknown Color";
        var pdfPath = CreatePdfDocument("test_unknown_color.pdf");
        var outputPath = CreateTestFilePath("test_unknown_color_output.pdf");

        _tool.Execute("add", text: watermarkText, path: pdfPath, outputPath: outputPath, color: "InvalidColor");

        Assert.True(File.Exists(outputPath));

        using var outputDoc = new Document(outputPath);
        var textAbsorber = new TextAbsorber();
        outputDoc.Pages.Accept(textAbsorber);
        Assert.Contains(watermarkText, textAbsorber.Text);
    }

    [Fact]
    public void Add_WithIsBackground_ShouldSetBackground()
    {
        const string watermarkText = "Background Watermark";
        var pdfPath = CreatePdfDocument("test_bg.pdf");
        var outputPath = CreateTestFilePath("test_bg_output.pdf");

        _tool.Execute("add", text: watermarkText, path: pdfPath, outputPath: outputPath, isBackground: true);

        Assert.True(File.Exists(outputPath));

        using var outputDoc = new Document(outputPath);
        var textAbsorber = new TextAbsorber();
        outputDoc.Pages.Accept(textAbsorber);
        Assert.Contains(watermarkText, textAbsorber.Text);
    }

    [Fact]
    public void Add_WithMultiplePages_ShouldApplyToAllPages()
    {
        const string watermarkText = "All Pages";
        var pdfPath = CreatePdfDocument("test_multi.pdf", 3);
        var outputPath = CreateTestFilePath("test_multi_output.pdf");

        var result = _tool.Execute("add", text: watermarkText, path: pdfPath, outputPath: outputPath);

        Assert.True(File.Exists(outputPath));
        Assert.StartsWith("Watermark added to 3 page(s)", result);

        using var outputDoc = new Document(outputPath);
        for (var i = 1; i <= 3; i++)
        {
            var textAbsorber = new TextAbsorber();
            outputDoc.Pages[i].Accept(textAbsorber);
            Assert.Contains(watermarkText, textAbsorber.Text);
        }
    }

    [Fact]
    public void Add_WithPageRange_ShouldApplyToSpecificPages()
    {
        const string watermarkText = "Selected Pages";
        var pdfPath = CreatePdfDocument("test_range.pdf", 3);
        var outputPath = CreateTestFilePath("test_range_output.pdf");

        var result = _tool.Execute("add", text: watermarkText, path: pdfPath, outputPath: outputPath,
            pageRange: "1,3");

        Assert.True(File.Exists(outputPath));
        Assert.StartsWith("Watermark added to 2 page(s)", result);

        using var outputDoc = new Document(outputPath);

        var textAbsorber1 = new TextAbsorber();
        outputDoc.Pages[1].Accept(textAbsorber1);
        Assert.Contains(watermarkText, textAbsorber1.Text);

        var textAbsorber3 = new TextAbsorber();
        outputDoc.Pages[3].Accept(textAbsorber3);
        Assert.Contains(watermarkText, textAbsorber3.Text);
    }

    [Fact]
    public void Add_WithPageRangeHyphen_ShouldApplyToRange()
    {
        const string watermarkText = "Range Pages";
        var pdfPath = CreatePdfDocument("test_hyphen.pdf", 4);
        var outputPath = CreateTestFilePath("test_hyphen_output.pdf");

        var result = _tool.Execute("add", text: watermarkText, path: pdfPath, outputPath: outputPath, pageRange: "2-4");

        Assert.True(File.Exists(outputPath));
        Assert.StartsWith("Watermark added to 3 page(s)", result);
    }

    [Fact]
    public void Add_WithMixedPageRange_ShouldApplyCorrectly()
    {
        const string watermarkText = "Mixed Range";
        var pdfPath = CreatePdfDocument("test_mixed.pdf", 4);
        var outputPath = CreateTestFilePath("test_mixed_output.pdf");

        var result = _tool.Execute("add", text: watermarkText, path: pdfPath, outputPath: outputPath,
            pageRange: "1,3-4");

        Assert.True(File.Exists(outputPath));
        Assert.StartsWith("Watermark added to 3 page(s)", result);
    }

    [Fact]
    public void Add_WithAllOptions_ShouldApplyAllOptions()
    {
        const string watermarkText = "Confidential";
        var pdfPath = CreatePdfDocument("test_all.pdf", 3);
        var outputPath = CreateTestFilePath("test_all_output.pdf");

        var result = _tool.Execute("add", text: watermarkText, path: pdfPath, outputPath: outputPath,
            fontName: "Arial", fontSize: 72, opacity: 0.3, rotation: 45, color: "Red",
            pageRange: "1-2", isBackground: true, horizontalAlignment: "Center", verticalAlignment: "Center");

        Assert.True(File.Exists(outputPath));
        Assert.StartsWith("Watermark added to 2 page(s)", result);

        using var outputDoc = new Document(outputPath);
        var textAbsorber = new TextAbsorber();
        outputDoc.Pages[1].Accept(textAbsorber);
        Assert.Contains(watermarkText, textAbsorber.Text);
    }

    [SkippableFact]
    public void Add_WithManyPages_ShouldApplyToAllPages()
    {
        SkipInEvaluationMode(AsposeLibraryType.Pdf, "5 pages exceeds 4-page limit in evaluation mode");
        const string watermarkText = "All Pages";
        var pdfPath = CreatePdfDocument("test_many.pdf", 5);
        var outputPath = CreateTestFilePath("test_many_output.pdf");

        var result = _tool.Execute("add", text: watermarkText, path: pdfPath, outputPath: outputPath);

        Assert.True(File.Exists(outputPath));
        Assert.StartsWith("Watermark added to 5 page(s)", result);
    }

    [SkippableFact]
    public void Add_WithLargePageRange_ShouldApplyCorrectly()
    {
        SkipInEvaluationMode(AsposeLibraryType.Pdf, "10 pages exceeds 4-page limit in evaluation mode");
        var pdfPath = CreatePdfDocument("test_large_range.pdf", 10);
        var outputPath = CreateTestFilePath("test_large_range_output.pdf");

        var result = _tool.Execute("add", text: "Range Pages", path: pdfPath, outputPath: outputPath, pageRange: "2-5");

        Assert.True(File.Exists(outputPath));
        Assert.StartsWith("Watermark added to 4 page(s)", result);
    }

    [Theory]
    [InlineData("ADD")]
    [InlineData("Add")]
    [InlineData("add")]
    public void Operation_ShouldBeCaseInsensitive_Add(string operation)
    {
        var pdfPath = CreatePdfDocument($"test_case_{operation}.pdf");
        var outputPath = CreateTestFilePath($"test_case_{operation}_output.pdf");

        var result = _tool.Execute(operation, text: "Watermark", path: pdfPath, outputPath: outputPath);

        Assert.StartsWith("Watermark added to 1 page(s)", result);
    }

    #endregion

    #region Exception

    [Fact]
    public void Execute_WithUnknownOperation_ShouldThrowArgumentException()
    {
        var pdfPath = CreatePdfDocument("test_unknown_op.pdf");
        var ex = Assert.Throws<ArgumentException>(() => _tool.Execute("unknown", text: "Test", path: pdfPath));
        Assert.StartsWith("Unknown operation: unknown", ex.Message);
    }

    [Fact]
    public void Add_WithMissingText_ShouldThrowArgumentException()
    {
        var pdfPath = CreatePdfDocument("test_missing_text.pdf");
        var ex = Assert.Throws<ArgumentException>(() => _tool.Execute("add", pdfPath, text: null));
        Assert.Equal("text is required for add operation", ex.Message);
    }

    [Fact]
    public void Add_WithEmptyText_ShouldThrowArgumentException()
    {
        var pdfPath = CreatePdfDocument("test_empty_text.pdf");
        var ex = Assert.Throws<ArgumentException>(() => _tool.Execute("add", pdfPath, text: ""));
        Assert.Equal("text is required for add operation", ex.Message);
    }

    [Fact]
    public void Add_WithInvalidPageRange_ShouldThrowArgumentException()
    {
        var pdfPath = CreatePdfDocument("test_invalid_range.pdf", 3);
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("add", text: "Test", path: pdfPath, pageRange: "invalid"));
        Assert.StartsWith("Invalid page number", ex.Message);
    }

    [Fact]
    public void Add_WithOutOfBoundsPageRange_ShouldThrowArgumentException()
    {
        var pdfPath = CreatePdfDocument("test_oob_range.pdf", 3);
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("add", text: "Test", path: pdfPath, pageRange: "1,5"));
        Assert.StartsWith("Page number 5 is out of bounds", ex.Message);
    }

    [Fact]
    public void Add_WithInvalidRangeFormat_ShouldThrowArgumentException()
    {
        var pdfPath = CreatePdfDocument("test_invalid_format.pdf", 3);
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("add", text: "Test", path: pdfPath, pageRange: "3-1"));
        Assert.StartsWith("Page range '3-1' is out of bounds", ex.Message);
    }

    [Fact]
    public void Execute_WithNoPathOrSessionId_ShouldThrowException()
    {
        Assert.ThrowsAny<Exception>(() => _tool.Execute("add", text: "Test", path: null, sessionId: null));
    }

    #endregion

    #region Session

    [Fact]
    public void Add_WithSessionId_ShouldModifyInMemory()
    {
        const string watermarkText = "Confidential";
        var pdfPath = CreatePdfDocument("test_session_add.pdf");
        var sessionId = OpenSession(pdfPath);

        var result = _tool.Execute("add", text: watermarkText, sessionId: sessionId);

        Assert.StartsWith("Watermark added to 1 page(s)", result);
        Assert.Contains(sessionId, result);
        var document = SessionManager.GetDocument<Document>(sessionId);
        Assert.NotNull(document);
        Assert.Single(document.Pages);

        var textAbsorber = new TextAbsorber();
        document.Pages.Accept(textAbsorber);
        Assert.Contains(watermarkText, textAbsorber.Text);
    }

    [Fact]
    public void Add_WithSessionId_AndOptions_ShouldApplyOptionsInMemory()
    {
        const string watermarkText = "DRAFT";
        var pdfPath = CreatePdfDocument("test_session_options.pdf", 2);
        var sessionId = OpenSession(pdfPath);

        var result = _tool.Execute("add", text: watermarkText, sessionId: sessionId,
            fontName: "Arial", fontSize: 72, opacity: 0.5, rotation: 45, color: "Red");

        Assert.StartsWith("Watermark added to 2 page(s)", result);
        var document = SessionManager.GetDocument<Document>(sessionId);
        Assert.Equal(2, document.Pages.Count);

        var textAbsorber = new TextAbsorber();
        document.Pages.Accept(textAbsorber);
        Assert.Contains(watermarkText, textAbsorber.Text);
    }

    [Fact]
    public void Add_WithSessionId_AndPageRange_ShouldApplyToSpecificPages()
    {
        const string watermarkText = "Selected";
        var pdfPath = CreatePdfDocument("test_session_range.pdf", 3);
        var sessionId = OpenSession(pdfPath);

        var result = _tool.Execute("add", text: watermarkText, sessionId: sessionId, pageRange: "1,3");

        Assert.StartsWith("Watermark added to 2 page(s)", result);
        var document = SessionManager.GetDocument<Document>(sessionId);
        Assert.Equal(3, document.Pages.Count);
    }

    [Fact]
    public void Execute_WithInvalidSessionId_ShouldThrowKeyNotFoundException()
    {
        Assert.Throws<KeyNotFoundException>(() => _tool.Execute("add", text: "Test", sessionId: "invalid_session"));
    }

    [Fact]
    public void Execute_WithBothPathAndSessionId_ShouldPreferSessionId()
    {
        var pdfPath1 = CreatePdfDocument("test_path_watermark.pdf");
        var pdfPath2 = CreatePdfDocument("test_session_watermark.pdf", 3);
        var sessionId = OpenSession(pdfPath2);

        var result = _tool.Execute("add", text: "Test", path: pdfPath1, sessionId: sessionId);

        Assert.StartsWith("Watermark added to 3 page(s)", result);
    }

    #endregion
}