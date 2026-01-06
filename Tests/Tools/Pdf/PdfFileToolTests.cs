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

    private string CreateTestPdf(string fileName, int pageCount = 1)
    {
        var filePath = CreateTestFilePath(fileName);
        using var document = new Document();
        for (var i = 0; i < pageCount; i++)
        {
            var page = document.Pages.Add();
            page.Paragraphs.Add(new TextFragment($"Page {i + 1}"));
        }

        document.Save(filePath);
        return filePath;
    }

    #region General

    [Fact]
    public void Create_ShouldCreateNewPdf()
    {
        var outputPath = CreateTestFilePath("test_create.pdf");
        var result = _tool.Execute("create", outputPath: outputPath);

        Assert.StartsWith("PDF document created", result);
        Assert.True(File.Exists(outputPath));

        using var doc = new Document(outputPath);
        Assert.True(doc.Pages.Count >= 0);
    }

    [Fact]
    public void Merge_WithTwoPdfs_ShouldMerge()
    {
        var pdf1Path = CreateTestPdf("test_merge1.pdf");
        var pdf2Path = CreateTestPdf("test_merge2.pdf");
        var outputPath = CreateTestFilePath("test_merge_output.pdf");

        var result = _tool.Execute("merge", outputPath: outputPath,
            inputPaths: [pdf1Path, pdf2Path]);

        Assert.StartsWith("Merged 2 PDF documents", result);
        using var document = new Document(outputPath);
        Assert.Equal(2, document.Pages.Count);
    }

    [Fact]
    public void Merge_WithThreePdfs_ShouldMergeAll()
    {
        var pdf1Path = CreateTestPdf("test_merge3_1.pdf");
        var pdf2Path = CreateTestPdf("test_merge3_2.pdf");
        var pdf3Path = CreateTestPdf("test_merge3_3.pdf");
        var outputPath = CreateTestFilePath("test_merge3_output.pdf");

        var result = _tool.Execute("merge", outputPath: outputPath,
            inputPaths: [pdf1Path, pdf2Path, pdf3Path]);

        Assert.StartsWith("Merged 3 PDF documents", result);
        using var document = new Document(outputPath);
        Assert.Equal(3, document.Pages.Count);
    }

    [Fact]
    public void Merge_WithSinglePdf_ShouldCreateOutput()
    {
        var pdfPath = CreateTestPdf("test_merge_single.pdf");
        var outputPath = CreateTestFilePath("test_merge_single_output.pdf");

        var result = _tool.Execute("merge", outputPath: outputPath, inputPaths: [pdfPath]);

        Assert.StartsWith("Merged 1 PDF document", result);
        Assert.True(File.Exists(outputPath));

        using var doc = new Document(outputPath);
        Assert.Single(doc.Pages);
    }

    [Fact]
    public void Split_ByPagesPerFile_ShouldSplitIntoMultipleFiles()
    {
        var pdfPath = CreateTestPdf("test_split.pdf", 2);
        var outputDir = Path.Combine(TestDir, "split_output");
        Directory.CreateDirectory(outputDir);

        var result = _tool.Execute("split", pdfPath, outputDir: outputDir, pagesPerFile: 1);

        Assert.StartsWith("PDF split into 2 files", result);
        var files = Directory.GetFiles(outputDir, "*.pdf");
        Assert.Equal(2, files.Length);

        foreach (var file in files)
        {
            using var doc = new Document(file);
            Assert.Single(doc.Pages);
        }
    }

    [Fact]
    public void Split_WithStartAndEndPage_ShouldExtractPageRange()
    {
        var pdfPath = CreateTestPdf("test_split_range.pdf", 3);
        var outputDir = Path.Combine(TestDir, "split_range_output");
        Directory.CreateDirectory(outputDir);

        _tool.Execute("split", pdfPath, outputDir: outputDir,
            startPage: 1, endPage: 2);

        var files = Directory.GetFiles(outputDir, "*.pdf");
        Assert.Single(files);
        Assert.Contains("pages_1-2", Path.GetFileName(files[0]));

        using var outputDoc = new Document(files[0]);
        Assert.Equal(2, outputDoc.Pages.Count);
    }

    [Fact]
    public void Split_WithStartPageOnly_ShouldExtractFromStartToEnd()
    {
        var pdfPath = CreateTestPdf("test_split_start.pdf", 3);
        var outputDir = Path.Combine(TestDir, "split_start_output");
        Directory.CreateDirectory(outputDir);

        _tool.Execute("split", pdfPath, outputDir: outputDir, startPage: 2);

        var files = Directory.GetFiles(outputDir, "*.pdf");
        Assert.Single(files);
        Assert.Contains("pages_2-3", Path.GetFileName(files[0]));

        using var outputDoc = new Document(files[0]);
        Assert.Equal(2, outputDoc.Pages.Count);
    }

    [Fact]
    public void Split_WithEndPageOnly_ShouldExtractFromBeginning()
    {
        var pdfPath = CreateTestPdf("test_split_end.pdf", 3);
        var outputDir = Path.Combine(TestDir, "split_end_output");
        Directory.CreateDirectory(outputDir);

        _tool.Execute("split", pdfPath, outputDir: outputDir, endPage: 2);

        var files = Directory.GetFiles(outputDir, "*.pdf");
        Assert.Single(files);
        Assert.Contains("pages_1-2", Path.GetFileName(files[0]));

        using var outputDoc = new Document(files[0]);
        Assert.Equal(2, outputDoc.Pages.Count);
    }

    [SkippableFact]
    public void Split_WithMultiplePagesPerFile_ShouldSplitCorrectly()
    {
        SkipInEvaluationMode(AsposeLibraryType.Pdf, "5 pages exceeds 4-page limit");
        var pdfPath = CreateTestPdf("test_split_multi.pdf", 5);
        var outputDir = Path.Combine(TestDir, "split_multi_output");
        Directory.CreateDirectory(outputDir);

        var result = _tool.Execute("split", pdfPath, outputDir: outputDir, pagesPerFile: 2);

        Assert.StartsWith("PDF split into 3 files", result);
        var files = Directory.GetFiles(outputDir, "*.pdf");
        Assert.Equal(3, files.Length);
    }

    [Fact]
    public void Compress_WithAllOptions_ShouldCompress()
    {
        var pdfPath = CreateTestPdf("test_compress.pdf");
        var outputPath = CreateTestFilePath("test_compress_output.pdf");

        var result = _tool.Execute("compress", pdfPath, outputPath: outputPath,
            compressImages: true, compressFonts: true, removeUnusedObjects: true);

        Assert.True(File.Exists(outputPath));

        Assert.StartsWith("PDF compressed", result);

        using var doc = new Document(outputPath);
        Assert.True(doc.Pages.Count > 0);
    }

    [Fact]
    public void Compress_WithNoCompression_ShouldStillCreateOutput()
    {
        var pdfPath = CreateTestPdf("test_compress_none.pdf");
        var outputPath = CreateTestFilePath("test_compress_none_output.pdf");
        _tool.Execute("compress", pdfPath, outputPath: outputPath,
            compressImages: false, compressFonts: false, removeUnusedObjects: false);
        Assert.True(File.Exists(outputPath));
    }

    [Fact]
    public void Encrypt_ShouldEncryptPdf()
    {
        var pdfPath = CreateTestPdf("test_encrypt.pdf");
        var outputPath = CreateTestFilePath("test_encrypt_output.pdf");
        var result = _tool.Execute("encrypt", pdfPath, outputPath: outputPath,
            userPassword: "user123", ownerPassword: "owner123");

        Assert.StartsWith("PDF encrypted", result);
        Assert.True(File.Exists(outputPath));

        Assert.ThrowsAny<Exception>(() =>
        {
            using var doc = new Document(outputPath);
            _ = doc.Pages.Count;
        });

        using var decryptedDoc = new Document(outputPath, "user123");
        Assert.True(decryptedDoc.Pages.Count > 0);
    }

    [Fact]
    public void Linearize_ShouldOptimizeForFastWebView()
    {
        var pdfPath = CreateTestPdf("test_linearize.pdf");
        var outputPath = CreateTestFilePath("test_linearize_output.pdf");

        var result = _tool.Execute("linearize", pdfPath, outputPath: outputPath);

        Assert.StartsWith("PDF linearized", result);
        Assert.True(File.Exists(outputPath));

        using var doc = new Document(outputPath);
        Assert.True(doc.Pages.Count > 0);
    }

    [Theory]
    [InlineData("CREATE")]
    [InlineData("Create")]
    [InlineData("create")]
    public void Operation_ShouldBeCaseInsensitive_Create(string operation)
    {
        var outputPath = CreateTestFilePath($"test_case_{operation}.pdf");
        var result = _tool.Execute(operation, outputPath: outputPath);
        Assert.StartsWith("PDF document created", result);
    }

    [Theory]
    [InlineData("COMPRESS")]
    [InlineData("Compress")]
    [InlineData("compress")]
    public void Operation_ShouldBeCaseInsensitive_Compress(string operation)
    {
        var pdfPath = CreateTestPdf($"test_case_{operation}.pdf");
        var outputPath = CreateTestFilePath($"test_case_{operation}_output.pdf");
        var result = _tool.Execute(operation, pdfPath, outputPath: outputPath);
        Assert.StartsWith("PDF compressed", result);
    }

    #endregion

    #region Exception

    [Fact]
    public void Execute_WithUnknownOperation_ShouldThrowArgumentException()
    {
        var ex = Assert.Throws<ArgumentException>(() => _tool.Execute("unknown"));
        Assert.StartsWith("Unknown operation: unknown", ex.Message);
    }

    [Fact]
    public void Create_WithMissingOutputPath_ShouldThrowArgumentException()
    {
        var ex = Assert.Throws<ArgumentException>(() => _tool.Execute("create"));
        Assert.Equal("outputPath is required for create operation", ex.Message);
    }

    [Fact]
    public void Merge_WithMissingOutputPath_ShouldThrowArgumentException()
    {
        var pdfPath = CreateTestPdf("test_merge_no_output.pdf");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("merge", inputPaths: [pdfPath]));
        Assert.Equal("outputPath is required for merge operation", ex.Message);
    }

    [Fact]
    public void Merge_WithEmptyInputPaths_ShouldThrowArgumentException()
    {
        var outputPath = CreateTestFilePath("test_merge_empty.pdf");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("merge", outputPath: outputPath, inputPaths: []));
        Assert.Equal("inputPaths is required for merge operation", ex.Message);
    }

    [Fact]
    public void Split_WithMissingOutputDir_ShouldThrowArgumentException()
    {
        var pdfPath = CreateTestPdf("test_split_no_dir.pdf");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("split", pdfPath));
        Assert.Equal("outputDir is required for split operation", ex.Message);
    }

    [Theory]
    [InlineData(0)]
    [InlineData(-1)]
    [InlineData(1001)]
    public void Split_WithInvalidPagesPerFile_ShouldThrowArgumentException(int pagesPerFile)
    {
        var pdfPath = CreateTestPdf("test_split_invalid_pages.pdf");
        var outputDir = Path.Combine(TestDir, "split_invalid_pages");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("split", pdfPath, outputDir: outputDir, pagesPerFile: pagesPerFile));
        Assert.Equal("pagesPerFile must be between 1 and 1000", ex.Message);
    }

    [Fact]
    public void Split_WithInvalidStartPage_ShouldThrowArgumentException()
    {
        var pdfPath = CreateTestPdf("test_split_invalid_start.pdf");
        var outputDir = Path.Combine(TestDir, "split_invalid_start");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("split", pdfPath, outputDir: outputDir, startPage: 0));
        Assert.StartsWith("startPage must be between 1 and", ex.Message);
    }

    [Fact]
    public void Split_WithInvalidEndPage_ShouldThrowArgumentException()
    {
        var pdfPath = CreateTestPdf("test_split_invalid_end.pdf");
        var outputDir = Path.Combine(TestDir, "split_invalid_end");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("split", pdfPath, outputDir: outputDir, startPage: 1, endPage: 100));
        Assert.StartsWith("endPage must be between", ex.Message);
    }

    [Fact]
    public void Split_WithEndPageLessThanStartPage_ShouldThrowArgumentException()
    {
        var pdfPath = CreateTestPdf("test_split_end_less.pdf", 3);
        var outputDir = Path.Combine(TestDir, "split_end_less");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("split", pdfPath, outputDir: outputDir, startPage: 3, endPage: 1));
        Assert.StartsWith("endPage must be between", ex.Message);
    }

    [Fact]
    public void Encrypt_WithMissingUserPassword_ShouldThrowArgumentException()
    {
        var pdfPath = CreateTestPdf("test_encrypt_no_user.pdf");
        var outputPath = CreateTestFilePath("test_encrypt_no_user_output.pdf");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("encrypt", pdfPath, outputPath: outputPath, ownerPassword: "owner"));
        Assert.Equal("userPassword is required for encrypt operation", ex.Message);
    }

    [Fact]
    public void Encrypt_WithMissingOwnerPassword_ShouldThrowArgumentException()
    {
        var pdfPath = CreateTestPdf("test_encrypt_no_owner.pdf");
        var outputPath = CreateTestFilePath("test_encrypt_no_owner_output.pdf");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("encrypt", pdfPath, outputPath: outputPath, userPassword: "user"));
        Assert.Equal("ownerPassword is required for encrypt operation", ex.Message);
    }

    [Fact]
    public void Compress_WithNoPathAndNoSessionId_ShouldThrowException()
    {
        var outputPath = CreateTestFilePath("test_compress_no_path.pdf");
        Assert.ThrowsAny<Exception>(() =>
            _tool.Execute("compress", outputPath: outputPath));
    }

    #endregion

    #region Session

    [Fact]
    public void Compress_WithSessionId_ShouldCompressInSession()
    {
        var pdfPath = CreateTestPdf("test_session_compress.pdf");
        var sessionId = OpenSession(pdfPath);

        var result = _tool.Execute("compress", sessionId: sessionId, compressImages: true);

        Assert.StartsWith("PDF compressed", result);
        Assert.Contains(sessionId, result);
    }

    [Fact]
    public void Linearize_WithSessionId_ShouldLinearizeInSession()
    {
        var pdfPath = CreateTestPdf("test_session_linearize.pdf");
        var sessionId = OpenSession(pdfPath);

        var result = _tool.Execute("linearize", sessionId: sessionId);

        Assert.StartsWith("PDF linearized", result);
        Assert.Contains(sessionId, result);
    }

    [Fact]
    public void Encrypt_WithSessionId_ShouldEncryptInSession()
    {
        var pdfPath = CreateTestPdf("test_session_encrypt.pdf");
        var sessionId = OpenSession(pdfPath);

        var result = _tool.Execute("encrypt", sessionId: sessionId,
            userPassword: "user", ownerPassword: "owner");

        Assert.StartsWith("PDF encrypted", result);
        Assert.Contains(sessionId, result);
    }

    [Fact]
    public void Split_WithSessionId_ShouldSplitFromSession()
    {
        var pdfPath = CreateTestPdf("test_session_split.pdf", 2);
        var sessionId = OpenSession(pdfPath);
        var outputDir = Path.Combine(TestDir, "session_split_output");
        Directory.CreateDirectory(outputDir);

        var result = _tool.Execute("split", sessionId: sessionId,
            outputDir: outputDir, pagesPerFile: 1);

        Assert.StartsWith("PDF split into 2 files", result);
        var files = Directory.GetFiles(outputDir, "*.pdf");
        Assert.Equal(2, files.Length);
    }

    [Fact]
    public void Compress_WithSessionId_ShouldModifyInMemory()
    {
        var pdfPath = CreateTestPdf("test_session_memory.pdf");
        var sessionId = OpenSession(pdfPath);

        _tool.Execute("compress", sessionId: sessionId, compressImages: true);

        var document = SessionManager.GetDocument<Document>(sessionId);
        Assert.NotNull(document);
        Assert.True(document.Pages.Count > 0);
    }

    [Fact]
    public void Execute_WithInvalidSessionId_ShouldThrowKeyNotFoundException()
    {
        Assert.Throws<KeyNotFoundException>(() =>
            _tool.Execute("compress", sessionId: "invalid_session"));
    }

    [Fact]
    public void Execute_WithBothPathAndSessionId_ShouldPreferSessionId()
    {
        var pdfPath1 = CreateTestPdf("test_path_file.pdf");
        var pdfPath2 = CreateTestPdf("test_session_file.pdf", 3);
        var sessionId = OpenSession(pdfPath2);
        var outputDir = Path.Combine(TestDir, "prefer_session_output");
        Directory.CreateDirectory(outputDir);

        var result = _tool.Execute("split", pdfPath1, sessionId, outputDir: outputDir, pagesPerFile: 1);

        Assert.StartsWith("PDF split into 3 files", result);
    }

    #endregion
}