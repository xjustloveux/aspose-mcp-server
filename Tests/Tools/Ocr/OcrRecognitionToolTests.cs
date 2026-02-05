using Aspose.Pdf;
using Aspose.Pdf.Text;
using AsposeMcpServer.Results.Ocr;
using AsposeMcpServer.Tests.Infrastructure;
using AsposeMcpServer.Tools.Ocr;

namespace AsposeMcpServer.Tests.Tools.Ocr;

/// <summary>
///     Unit tests for OcrRecognitionTool.
///     Tests tool construction, operation routing, and parameter validation.
/// </summary>
public class OcrRecognitionToolTests : TestBase
{
    private readonly OcrRecognitionTool _tool = new();

    #region Construction

    [Fact]
    public void Constructor_ShouldCreateToolWithHandlerRegistry()
    {
        var tool = new OcrRecognitionTool();

        Assert.NotNull(tool);
    }

    #endregion

    /// <summary>
    ///     Creates a simple PDF file with text content for OCR testing.
    /// </summary>
    /// <param name="text">The text content to add to the PDF.</param>
    /// <returns>The path to the created PDF file.</returns>
    private string CreatePdfWithText(string text)
    {
        var pdfPath = CreateTestFilePath($"ocr_test_{Guid.NewGuid()}.pdf");
        var doc = new Document();
        var page = doc.Pages.Add();
        page.Paragraphs.Add(new TextFragment(text));
        doc.Save(pdfPath);
        doc.Dispose();
        return pdfPath;
    }

    /// <summary>
    ///     Creates a simple BMP image file for testing.
    /// </summary>
    /// <returns>The full path to the created image file.</returns>
    private string CreateTestImageFile()
    {
        var width = 10;
        var height = 10;
        var bmp = new byte[width * height * 3 + 54];
        bmp[0] = 0x42;
        bmp[1] = 0x4D;
        var fileSize = bmp.Length;
        bmp[2] = (byte)(fileSize & 0xFF);
        bmp[3] = (byte)((fileSize >> 8) & 0xFF);
        bmp[4] = (byte)((fileSize >> 16) & 0xFF);
        bmp[5] = (byte)((fileSize >> 24) & 0xFF);
        bmp[10] = 54;
        bmp[14] = 40;
        bmp[18] = (byte)(width & 0xFF);
        bmp[19] = (byte)((width >> 8) & 0xFF);
        bmp[22] = (byte)(height & 0xFF);
        bmp[23] = (byte)((height >> 8) & 0xFF);
        bmp[26] = 1;
        bmp[28] = 24;
        for (var i = 54; i < bmp.Length; i += 3)
        {
            bmp[i] = 255;
            bmp[i + 1] = 0;
            bmp[i + 2] = 0;
        }

        var filePath = CreateTestFilePath($"test_image_{Guid.NewGuid()}.bmp");
        File.WriteAllBytes(filePath, bmp);
        return filePath;
    }

    #region Operation Routing

    [Fact]
    public void Execute_WithUnknownOperation_ShouldThrowArgumentException()
    {
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("unknown_operation", Path.Combine(TestDir, "test.png")));

        Assert.Contains("Unknown operation", ex.Message);
    }

    [Theory]
    [InlineData("RECOGNIZE")]
    [InlineData("Recognize")]
    [InlineData("recognize")]
    public void Execute_OperationIsCaseInsensitive_Recognize(string operation)
    {
        var tempFile = Path.Combine(TestDir, "nonexistent.png");

        Assert.Throws<FileNotFoundException>(() =>
            _tool.Execute(operation, tempFile));
    }

    [Theory]
    [InlineData("RECOGNIZE_PDF")]
    [InlineData("Recognize_Pdf")]
    [InlineData("recognize_pdf")]
    public void Execute_OperationIsCaseInsensitive_RecognizePdf(string operation)
    {
        var tempFile = CreateTestFilePath("nonexistent.pdf");

        Assert.Throws<FileNotFoundException>(() =>
            _tool.Execute(operation, tempFile, CreateTestFilePath("output.docx"),
                targetFormat: "docx"));
    }

    [Theory]
    [InlineData("RECOGNIZE_RECEIPT")]
    [InlineData("Recognize_Receipt")]
    [InlineData("recognize_receipt")]
    public void Execute_OperationIsCaseInsensitive_RecognizeReceipt(string operation)
    {
        var tempFile = Path.Combine(TestDir, "nonexistent_receipt.jpg");

        Assert.Throws<FileNotFoundException>(() =>
            _tool.Execute(operation, tempFile));
    }

    [Theory]
    [InlineData("RECOGNIZE_ID_CARD")]
    [InlineData("Recognize_Id_Card")]
    [InlineData("recognize_id_card")]
    public void Execute_OperationIsCaseInsensitive_RecognizeIdCard(string operation)
    {
        var tempFile = Path.Combine(TestDir, "nonexistent_id_card.jpg");

        Assert.Throws<FileNotFoundException>(() =>
            _tool.Execute(operation, tempFile));
    }

    [Theory]
    [InlineData("RECOGNIZE_PASSPORT")]
    [InlineData("Recognize_Passport")]
    [InlineData("recognize_passport")]
    public void Execute_OperationIsCaseInsensitive_RecognizePassport(string operation)
    {
        var tempFile = Path.Combine(TestDir, "nonexistent_passport.jpg");

        Assert.Throws<FileNotFoundException>(() =>
            _tool.Execute(operation, tempFile));
    }

    #endregion

    #region Error Handling

    [Fact]
    public void Execute_Recognize_WithNonexistentFile_ShouldThrowFileNotFoundException()
    {
        Assert.Throws<FileNotFoundException>(() =>
            _tool.Execute("recognize", Path.Combine(TestDir, "nonexistent.png")));
    }

    [Fact]
    public void Execute_Recognize_WithUnsupportedFormat_ShouldThrowArgumentException()
    {
        var tempFile = CreateTestFilePath("test.xyz");
        File.WriteAllText(tempFile, "dummy");

        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("recognize", tempFile));

        Assert.Contains("Unsupported file format", ex.Message);
    }

    [Fact]
    public void Execute_RecognizePdf_WithNonPdfInput_ShouldThrowArgumentException()
    {
        var tempFile = CreateTestFilePath("test.png");
        File.WriteAllText(tempFile, "dummy");

        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("recognize_pdf", tempFile, CreateTestFilePath("output.docx"),
                targetFormat: "docx"));

        Assert.Contains("PDF", ex.Message);
    }

    [Fact]
    public void Execute_RecognizePdf_WithUnsupportedTargetFormat_ShouldThrowArgumentException()
    {
        var tempFile = CreateTestFilePath("test.pdf");
        File.WriteAllText(tempFile, "dummy");

        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("recognize_pdf", tempFile, CreateTestFilePath("output.xyz"),
                targetFormat: "xyz"));

        Assert.Contains("Unsupported target format", ex.Message);
    }

    [Fact]
    public void Execute_RecognizeReceipt_WithNonexistentFile_ShouldThrowFileNotFoundException()
    {
        Assert.Throws<FileNotFoundException>(() =>
            _tool.Execute("recognize_receipt", Path.Combine(TestDir, "nonexistent_receipt.jpg")));
    }

    [Fact]
    public void Execute_RecognizeIdCard_WithNonexistentFile_ShouldThrowFileNotFoundException()
    {
        Assert.Throws<FileNotFoundException>(() =>
            _tool.Execute("recognize_id_card", Path.Combine(TestDir, "nonexistent_id_card.jpg")));
    }

    [Fact]
    public void Execute_RecognizePassport_WithNonexistentFile_ShouldThrowFileNotFoundException()
    {
        Assert.Throws<FileNotFoundException>(() =>
            _tool.Execute("recognize_passport", Path.Combine(TestDir, "nonexistent_passport.jpg")));
    }

    #endregion

    #region Happy Path

    [SkippableFact]
    public void Execute_Recognize_WithValidPdf_ShouldReturnRecognitionResult()
    {
        SkipInEvaluationMode(AsposeLibraryType.Ocr, "OCR recognition requires a valid license");
        var pdfPath = CreatePdfWithText("Tool-level OCR recognition test");

        var result = _tool.Execute("recognize", pdfPath);

        var data = GetResultData<OcrRecognitionResult>(result);
        Assert.NotNull(data.Text);
        Assert.True(data.PageCount >= 0);
        Assert.NotNull(data.Pages);
    }

    [SkippableFact]
    public void Execute_RecognizePdf_WithValidPdf_ShouldReturnConversionResult()
    {
        SkipInEvaluationMode(AsposeLibraryType.Ocr, "OCR conversion requires a valid license");
        var pdfPath = CreatePdfWithText("Tool-level OCR conversion test");
        var outputPath = CreateTestFilePath("tool_output.docx");

        var result = _tool.Execute("recognize_pdf", pdfPath, outputPath, targetFormat: "docx");

        var data = GetResultData<OcrConversionResult>(result);
        Assert.Equal(pdfPath, data.SourcePath);
        Assert.Equal(outputPath, data.OutputPath);
        Assert.Equal("docx", data.TargetFormat);
        Assert.True(data.PageCount >= 0);
        Assert.True(File.Exists(outputPath));
    }

    [SkippableFact]
    public void Execute_RecognizeReceipt_WithValidImage_ShouldReturnRecognitionResult()
    {
        SkipInEvaluationMode(AsposeLibraryType.Ocr, "OCR recognition requires a valid license");
        var tempImage = CreateTestImageFile();

        var result = _tool.Execute("recognize_receipt", tempImage);

        var data = GetResultData<OcrRecognitionResult>(result);
        Assert.NotNull(data.Text);
        Assert.NotNull(data.Pages);
    }

    [SkippableFact]
    public void Execute_RecognizeIdCard_WithValidImage_ShouldReturnRecognitionResult()
    {
        SkipInEvaluationMode(AsposeLibraryType.Ocr, "OCR recognition requires a valid license");
        var tempImage = CreateTestImageFile();

        var result = _tool.Execute("recognize_id_card", tempImage);

        var data = GetResultData<OcrRecognitionResult>(result);
        Assert.NotNull(data.Text);
        Assert.NotNull(data.Pages);
    }

    [SkippableFact]
    public void Execute_RecognizePassport_WithValidImage_ShouldReturnRecognitionResult()
    {
        SkipInEvaluationMode(AsposeLibraryType.Ocr, "OCR recognition requires a valid license");
        var tempImage = CreateTestImageFile();

        var result = _tool.Execute("recognize_passport", tempImage);

        var data = GetResultData<OcrRecognitionResult>(result);
        Assert.NotNull(data.Text);
        Assert.NotNull(data.Pages);
    }

    #endregion
}
