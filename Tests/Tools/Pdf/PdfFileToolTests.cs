using Aspose.Pdf;
using Aspose.Pdf.Text;
using AsposeMcpServer.Tests.Helpers;
using AsposeMcpServer.Tools.Pdf;

namespace AsposeMcpServer.Tests.Tools.Pdf;

public class PdfFileToolTests : PdfTestBase
{
    private readonly PdfFileTool _tool;

    public PdfFileToolTests()
    {
        _tool = new PdfFileTool(SessionManager);
    }

    private string CreateTestPdf(string fileName)
    {
        var filePath = CreateTestFilePath(fileName);
        var document = new Document();
        var page = document.Pages.Add();
        page.Paragraphs.Add(new TextFragment("Test PDF"));
        document.Save(filePath);
        return filePath;
    }

    #region General Tests

    [Fact]
    public void CreatePdf_ShouldCreateNewPdf()
    {
        var outputPath = CreateTestFilePath("test_create_pdf.pdf");
        _tool.Execute("create", outputPath: outputPath);
        Assert.True(File.Exists(outputPath), "PDF should be created");
    }

    [Fact]
    public void MergePdfs_ShouldMergeMultiplePdfs()
    {
        var pdf1Path = CreateTestPdf("test_merge1.pdf");
        var pdf2Path = CreateTestPdf("test_merge2.pdf");
        var outputPath = CreateTestFilePath("test_merge_output.pdf");
        _tool.Execute(
            "merge",
            outputPath: outputPath,
            inputPaths: [pdf1Path, pdf2Path]);
        Assert.True(File.Exists(outputPath), "Merged PDF should be created");
        var document = new Document(outputPath);
        Assert.True(document.Pages.Count >= 2, "Merged PDF should have multiple pages");
    }

    [Fact]
    public void SplitPdf_ShouldSplitIntoMultipleFiles()
    {
        var pdfPath = CreateTestPdf("test_split.pdf");
        var document = new Document(pdfPath);
        document.Pages.Add();
        document.Save(pdfPath);

        var outputDir = Path.Combine(TestDir, "split_output");
        Directory.CreateDirectory(outputDir);
        _tool.Execute(
            "split",
            pdfPath,
            outputDir: outputDir,
            pagesPerFile: 1);
        var files = Directory.GetFiles(outputDir);
        Assert.True(files.Length >= 2, "Should create multiple files for split pages");
    }

    [Fact]
    public void CompressPdf_ShouldCompressPdf()
    {
        var pdfPath = CreateTestPdf("test_compress.pdf");
        var outputPath = CreateTestFilePath("test_compress_output.pdf");
        _tool.Execute(
            "compress",
            pdfPath,
            outputPath: outputPath,
            compressImages: true);
        Assert.True(File.Exists(outputPath), "Compressed PDF should be created");
    }

    [Fact]
    public void EncryptPdf_ShouldEncryptPdf()
    {
        var pdfPath = CreateTestPdf("test_encrypt.pdf");
        var outputPath = CreateTestFilePath("test_encrypt_output.pdf");
        _tool.Execute(
            "encrypt",
            pdfPath,
            outputPath: outputPath,
            userPassword: "user123",
            ownerPassword: "owner123");
        Assert.True(File.Exists(outputPath), "Encrypted PDF should be created");
    }

    [Fact]
    public void CompressPdf_WithCompressFonts_ShouldCompressWithFontSubsetting()
    {
        var pdfPath = CreateTestPdf("test_compress_fonts.pdf");
        var outputPath = CreateTestFilePath("test_compress_fonts_output.pdf");
        _tool.Execute(
            "compress",
            pdfPath,
            outputPath: outputPath,
            compressImages: true,
            compressFonts: true);
        Assert.True(File.Exists(outputPath), "Compressed PDF with font subsetting should be created");
    }

    [Fact]
    public void CompressPdf_WithRemoveUnusedObjects_ShouldRemoveUnused()
    {
        var pdfPath = CreateTestPdf("test_compress_unused.pdf");
        var outputPath = CreateTestFilePath("test_compress_unused_output.pdf");
        _tool.Execute(
            "compress",
            pdfPath,
            outputPath: outputPath,
            removeUnusedObjects: true);
        Assert.True(File.Exists(outputPath), "Compressed PDF with unused objects removed should be created");
    }

    [Fact]
    public void CompressPdf_WithAllOptions_ShouldApplyAllCompression()
    {
        var pdfPath = CreateTestPdf("test_compress_all.pdf");
        var outputPath = CreateTestFilePath("test_compress_all_output.pdf");
        var result = _tool.Execute(
            "compress",
            pdfPath,
            outputPath: outputPath,
            compressImages: true,
            compressFonts: true,
            removeUnusedObjects: true);
        Assert.True(File.Exists(outputPath), "Fully compressed PDF should be created");
        Assert.Contains("compressed", result, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void CompressPdf_WithNoCompression_ShouldStillCreateOutput()
    {
        var pdfPath = CreateTestPdf("test_compress_none.pdf");
        var outputPath = CreateTestFilePath("test_compress_none_output.pdf");
        _tool.Execute(
            "compress",
            pdfPath,
            outputPath: outputPath,
            compressImages: false,
            compressFonts: false,
            removeUnusedObjects: false);
        Assert.True(File.Exists(outputPath), "PDF should be created even with no compression");
    }

    [SkippableFact]
    public void SplitPdf_WithMultiplePagesPerFile_ShouldSplitCorrectly()
    {
        // Skip in evaluation mode - Aspose.PDF limits collections to 4 elements
        SkipInEvaluationMode(AsposeLibraryType.Pdf, "This test requires more pages than allowed in evaluation mode");
        var pdfPath = CreateTestPdf("test_split_multi.pdf");

        var document = new Document(pdfPath);
        // Add multiple pages
        document.Pages.Add();
        document.Pages.Add();
        document.Pages.Add();
        document.Pages.Add(); // Total 5 pages
        document.Save(pdfPath);

        var outputDir = Path.Combine(TestDir, "split_multi_output");
        Directory.CreateDirectory(outputDir);
        _tool.Execute(
            "split",
            pdfPath,
            outputDir: outputDir,
            pagesPerFile: 2);
        var files = Directory.GetFiles(outputDir, "*.pdf");
        Assert.True(files.Length >= 2, "Should create multiple files when splitting with 2 pages per file");
    }

    [Fact]
    public void MergePdfs_WithThreePdfs_ShouldMergeAll()
    {
        var pdf1Path = CreateTestPdf("test_merge3_1.pdf");
        var pdf2Path = CreateTestPdf("test_merge3_2.pdf");
        var pdf3Path = CreateTestPdf("test_merge3_3.pdf");
        var outputPath = CreateTestFilePath("test_merge3_output.pdf");
        _tool.Execute(
            "merge",
            outputPath: outputPath,
            inputPaths: [pdf1Path, pdf2Path, pdf3Path]);
        Assert.True(File.Exists(outputPath), "Merged PDF should be created");
        var document = new Document(outputPath);
        Assert.True(document.Pages.Count >= 3, "Merged PDF should have at least 3 pages");
    }

    [Fact]
    public void Execute_WithUnknownOperation_ShouldThrowArgumentException()
    {
        var exception = Assert.Throws<ArgumentException>(() => _tool.Execute("unknown"));
        Assert.Contains("Unknown operation", exception.Message);
    }

    [Theory]
    [InlineData(0)]
    [InlineData(-1)]
    [InlineData(1001)]
    public void SplitPdf_WithInvalidPagesPerFile_ShouldThrowArgumentException(int pagesPerFile)
    {
        var pdfPath = CreateTestPdf("test_split_invalid.pdf");
        var outputDir = Path.Combine(TestDir, "split_invalid_output");
        var exception = Assert.Throws<ArgumentException>(() => _tool.Execute(
            "split",
            pdfPath,
            outputDir: outputDir,
            pagesPerFile: pagesPerFile));
        Assert.Contains("pagesPerFile must be between 1 and 1000", exception.Message);
    }

    [Fact]
    public void MergePdfs_WithEmptyInputPaths_ShouldThrowArgumentException()
    {
        var outputPath = CreateTestFilePath("test_merge_empty_output.pdf");
        var exception = Assert.Throws<ArgumentException>(() => _tool.Execute(
            "merge",
            outputPath: outputPath,
            inputPaths: Array.Empty<string>()));
        Assert.Contains("inputPaths is required", exception.Message);
    }

    [Fact]
    public void MergePdfs_WithSinglePdf_ShouldCreateOutput()
    {
        var pdfPath = CreateTestPdf("test_merge_single.pdf");
        var outputPath = CreateTestFilePath("test_merge_single_output.pdf");
        _tool.Execute(
            "merge",
            outputPath: outputPath,
            inputPaths: [pdfPath]);
        Assert.True(File.Exists(outputPath), "Merged PDF should be created even with single input");
    }

    [Fact]
    public void LinearizePdf_ShouldOptimizeForFastWebView()
    {
        var pdfPath = CreateTestPdf("test_linearize.pdf");
        var outputPath = CreateTestFilePath("test_linearize_output.pdf");
        var result = _tool.Execute(
            "linearize",
            pdfPath,
            outputPath: outputPath);
        Assert.True(File.Exists(outputPath), "Linearized PDF should be created");
        Assert.Contains("linearized", result, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void SplitPdf_WithStartAndEndPage_ShouldExtractPageRange()
    {
        var pdfPath = CreateTestPdf("test_split_range.pdf");
        using var document = new Document(pdfPath);
        document.Pages.Add();
        document.Pages.Add();
        document.Save(pdfPath);

        var outputDir = Path.Combine(TestDir, "split_range_output");
        Directory.CreateDirectory(outputDir);
        var result = _tool.Execute(
            "split",
            pdfPath,
            outputDir: outputDir,
            startPage: 1,
            endPage: 2);
        var files = Directory.GetFiles(outputDir, "*.pdf");
        Assert.Single(files);
        Assert.Contains("pages_1-2", result);
        using var outputDoc = new Document(files[0]);
        Assert.Equal(2, outputDoc.Pages.Count);
    }

    [Fact]
    public void SplitPdf_WithStartPageOnly_ShouldExtractFromStartToEnd()
    {
        var pdfPath = CreateTestPdf("test_split_start.pdf");
        using var document = new Document(pdfPath);
        document.Pages.Add();
        document.Pages.Add();
        document.Save(pdfPath);

        var outputDir = Path.Combine(TestDir, "split_start_output");
        Directory.CreateDirectory(outputDir);
        var result = _tool.Execute(
            "split",
            pdfPath,
            outputDir: outputDir,
            startPage: 2);
        var files = Directory.GetFiles(outputDir, "*.pdf");
        Assert.Single(files);
        Assert.Contains("pages_2-3", result);
    }

    [Fact]
    public void SplitPdf_WithEndPageOnly_ShouldExtractFromBeginning()
    {
        var pdfPath = CreateTestPdf("test_split_end.pdf");
        using var document = new Document(pdfPath);
        document.Pages.Add();
        document.Pages.Add();
        document.Save(pdfPath);

        var outputDir = Path.Combine(TestDir, "split_end_output");
        Directory.CreateDirectory(outputDir);
        var result = _tool.Execute(
            "split",
            pdfPath,
            outputDir: outputDir,
            endPage: 2);
        var files = Directory.GetFiles(outputDir, "*.pdf");
        Assert.Single(files);
        Assert.Contains("pages_1-2", result);
    }

    [Fact]
    public void SplitPdf_WithInvalidStartPage_ShouldThrowArgumentException()
    {
        var pdfPath = CreateTestPdf("test_split_invalid_start.pdf");
        var outputDir = Path.Combine(TestDir, "split_invalid_start_output");
        var exception = Assert.Throws<ArgumentException>(() => _tool.Execute(
            "split",
            pdfPath,
            outputDir: outputDir,
            startPage: 0));
        Assert.Contains("startPage must be between 1 and", exception.Message);
    }

    [Fact]
    public void SplitPdf_WithInvalidEndPage_ShouldThrowArgumentException()
    {
        var pdfPath = CreateTestPdf("test_split_invalid_end.pdf");
        var outputDir = Path.Combine(TestDir, "split_invalid_end_output");
        var exception = Assert.Throws<ArgumentException>(() => _tool.Execute(
            "split",
            pdfPath,
            outputDir: outputDir,
            startPage: 1,
            endPage: 100));
        Assert.Contains("endPage must be between", exception.Message);
    }

    [Fact]
    public void SplitPdf_WithEndPageLessThanStartPage_ShouldThrowArgumentException()
    {
        var pdfPath = CreateTestPdf("test_split_end_less.pdf");
        using var document = new Document(pdfPath);
        document.Pages.Add();
        document.Pages.Add();
        document.Save(pdfPath);

        var outputDir = Path.Combine(TestDir, "split_end_less_output");
        var exception = Assert.Throws<ArgumentException>(() => _tool.Execute(
            "split",
            pdfPath,
            outputDir: outputDir,
            startPage: 3,
            endPage: 1));
        Assert.Contains("endPage must be between", exception.Message);
    }

    #endregion

    #region Exception Tests

    [Fact]
    public void Execute_WithInvalidOperation_ShouldThrowArgumentException()
    {
        var exception = Assert.Throws<ArgumentException>(() => _tool.Execute("invalid_operation"));
        Assert.Contains("Unknown operation", exception.Message);
    }

    [Fact]
    public void MergePdfs_WithMissingOutputPath_ShouldThrowArgumentException()
    {
        var pdf1Path = CreateTestPdf("test_merge_no_output1.pdf");
        var pdf2Path = CreateTestPdf("test_merge_no_output2.pdf");
        var exception = Assert.Throws<ArgumentException>(() => _tool.Execute(
            "merge",
            inputPaths: [pdf1Path, pdf2Path]));
        Assert.Contains("outputPath is required", exception.Message);
    }

    #endregion

    #region Session ID Tests

    [Fact]
    public void CompressPdf_WithSessionId_ShouldCompressInSession()
    {
        var pdfPath = CreateTestPdf("test_session_compress.pdf");
        var sessionId = OpenSession(pdfPath);
        var result = _tool.Execute(
            "compress",
            sessionId: sessionId,
            compressImages: true);
        Assert.Contains("compressed", result, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("session", result);
    }

    [Fact]
    public void LinearizePdf_WithSessionId_ShouldLinearizeInSession()
    {
        var pdfPath = CreateTestPdf("test_session_linearize.pdf");
        var sessionId = OpenSession(pdfPath);
        var result = _tool.Execute(
            "linearize",
            sessionId: sessionId);
        Assert.Contains("linearized", result, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("session", result);
    }

    [Fact]
    public void CompressPdf_WithSessionId_ShouldModifyInMemory()
    {
        var pdfPath = CreateTestPdf("test_session_memory.pdf");
        var sessionId = OpenSession(pdfPath);
        _tool.Execute(
            "compress",
            sessionId: sessionId,
            compressImages: true,
            removeUnusedObjects: true);

        // Assert - verify in-memory document exists
        var document = SessionManager.GetDocument<Document>(sessionId);
        Assert.NotNull(document);
        Assert.True(document.Pages.Count > 0);
    }

    #endregion
}