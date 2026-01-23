using Aspose.Pdf;
using Aspose.Pdf.Text;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;
using AsposeMcpServer.Tools.Pdf;

namespace AsposeMcpServer.Tests.Tools.Pdf;

/// <summary>
///     Integration tests for PdfFileTool.
///     Focuses on session management, file I/O, and operation routing.
///     Detailed parameter validation and business logic tests are in Handler tests.
/// </summary>
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

    #region File I/O Smoke Tests

    [Fact]
    public void Create_ShouldCreateNewPdf()
    {
        var outputPath = CreateTestFilePath("test_create.pdf");
        var result = _tool.Execute("create", outputPath: outputPath);
        var data = GetResultData<SuccessResult>(result);
        Assert.StartsWith("PDF document created", data.Message);
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
        var data = GetResultData<SuccessResult>(result);
        Assert.StartsWith("Merged 2 PDF documents", data.Message);
        using var document = new Document(outputPath);
        Assert.Equal(2, document.Pages.Count);
    }

    [Fact]
    public void Split_ByPagesPerFile_ShouldSplitIntoMultipleFiles()
    {
        var pdfPath = CreateTestPdf("test_split.pdf", 2);
        var outputDir = Path.Combine(TestDir, "split_output");
        Directory.CreateDirectory(outputDir);

        var result = _tool.Execute("split", pdfPath, outputDir: outputDir, pagesPerFile: 1);
        var data = GetResultData<SuccessResult>(result);
        Assert.StartsWith("PDF split into 2 files", data.Message);
        var files = Directory.GetFiles(outputDir, "*.pdf");
        Assert.Equal(2, files.Length);

        foreach (var file in files)
        {
            using var doc = new Document(file);
            Assert.Single(doc.Pages);
        }
    }

    [Fact]
    public void Compress_WithAllOptions_ShouldCompress()
    {
        var pdfPath = CreateTestPdf("test_compress.pdf");
        var outputPath = CreateTestFilePath("test_compress_output.pdf");

        var result = _tool.Execute("compress", pdfPath, outputPath: outputPath,
            compressImages: true, compressFonts: true, removeUnusedObjects: true);

        Assert.True(File.Exists(outputPath));
        Assert.StartsWith("PDF compressed", GetResultData<string>(result));
    }

    [Fact]
    public void Encrypt_ShouldEncryptPdf()
    {
        var pdfPath = CreateTestPdf("test_encrypt.pdf");
        var outputPath = CreateTestFilePath("test_encrypt_output.pdf");
        var result = _tool.Execute("encrypt", pdfPath, outputPath: outputPath,
            userPassword: "user123", ownerPassword: "owner123");
        var data = GetResultData<SuccessResult>(result);
        Assert.StartsWith("PDF encrypted", data.Message);
        Assert.True(File.Exists(outputPath));

        using var decryptedDoc = new Document(outputPath, "user123");
        Assert.True(decryptedDoc.Pages.Count > 0);
    }

    [Fact]
    public void Linearize_ShouldOptimizeForFastWebView()
    {
        var pdfPath = CreateTestPdf("test_linearize.pdf");
        var outputPath = CreateTestFilePath("test_linearize_output.pdf");

        var result = _tool.Execute("linearize", pdfPath, outputPath: outputPath);
        Assert.StartsWith("PDF linearized", GetResultData<string>(result));
        Assert.True(File.Exists(outputPath));
    }

    #endregion

    #region Operation Routing

    [Theory]
    [InlineData("CREATE")]
    [InlineData("Create")]
    [InlineData("create")]
    public void Operation_ShouldBeCaseInsensitive(string operation)
    {
        var outputPath = CreateTestFilePath($"test_case_{operation}.pdf");
        var result = _tool.Execute(operation, outputPath: outputPath);
        var data = GetResultData<SuccessResult>(result);
        Assert.StartsWith("PDF document created", data.Message);
    }

    [Fact]
    public void Execute_WithUnknownOperation_ShouldThrowArgumentException()
    {
        var pdfPath = CreateTestPdf("test_unknown_op.pdf");
        var ex = Assert.Throws<ArgumentException>(() => _tool.Execute("unknown", pdfPath));
        Assert.StartsWith("Unknown operation: unknown", ex.Message);
    }

    [Fact]
    public void Execute_WithNoPathOrSessionId_ShouldThrowException()
    {
        var outputPath = CreateTestFilePath("test_compress_no_path.pdf");
        Assert.ThrowsAny<Exception>(() =>
            _tool.Execute("compress", outputPath: outputPath));
    }

    #endregion

    #region Session Management

    [Fact]
    public void Compress_WithSessionId_ShouldCompressInSession()
    {
        var pdfPath = CreateTestPdf("test_session_compress.pdf");
        var sessionId = OpenSession(pdfPath);

        var result = _tool.Execute("compress", sessionId: sessionId, compressImages: true);
        var data = GetResultData<SuccessResult>(result);
        Assert.StartsWith("PDF compressed", data.Message);
        var output = GetResultOutput<SuccessResult>(result);
        Assert.True(output.IsSession);
    }

    [Fact]
    public void Linearize_WithSessionId_ShouldLinearizeInSession()
    {
        var pdfPath = CreateTestPdf("test_session_linearize.pdf");
        var sessionId = OpenSession(pdfPath);

        var result = _tool.Execute("linearize", sessionId: sessionId);
        var data = GetResultData<SuccessResult>(result);
        Assert.StartsWith("PDF linearized", data.Message);
        var output = GetResultOutput<SuccessResult>(result);
        Assert.True(output.IsSession);
    }

    [Fact]
    public void Encrypt_WithSessionId_ShouldEncryptInSession()
    {
        var pdfPath = CreateTestPdf("test_session_encrypt.pdf");
        var sessionId = OpenSession(pdfPath);

        var result = _tool.Execute("encrypt", sessionId: sessionId,
            userPassword: "user", ownerPassword: "owner");
        var data = GetResultData<SuccessResult>(result);
        Assert.StartsWith("PDF encrypted", data.Message);
        var output = GetResultOutput<SuccessResult>(result);
        Assert.True(output.IsSession);
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
        var data = GetResultData<SuccessResult>(result);
        Assert.StartsWith("PDF split into 2 files", data.Message);
        var output = GetResultOutput<SuccessResult>(result);
        Assert.True(output.IsSession);
        var files = Directory.GetFiles(outputDir, "*.pdf");
        Assert.Equal(2, files.Length);
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
        var data = GetResultData<SuccessResult>(result);
        Assert.StartsWith("PDF split into 3 files", data.Message);
    }

    #endregion
}
