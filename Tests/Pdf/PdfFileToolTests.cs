using System.Text.Json.Nodes;
using Aspose.Pdf;
using Aspose.Pdf.Text;
using AsposeMcpServer.Tests.Helpers;
using AsposeMcpServer.Tools.Pdf;

namespace AsposeMcpServer.Tests.Pdf;

public class PdfFileToolTests : PdfTestBase
{
    private readonly PdfFileTool _tool = new();

    private string CreateTestPdf(string fileName)
    {
        var filePath = CreateTestFilePath(fileName);
        var document = new Document();
        var page = document.Pages.Add();
        page.Paragraphs.Add(new TextFragment("Test PDF"));
        document.Save(filePath);
        return filePath;
    }

    [Fact]
    public async Task CreatePdf_ShouldCreateNewPdf()
    {
        // Arrange
        var outputPath = CreateTestFilePath("test_create_pdf.pdf");
        var arguments = new JsonObject
        {
            ["operation"] = "create",
            ["outputPath"] = outputPath
        };

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.True(File.Exists(outputPath), "PDF should be created");
    }

    [Fact]
    public async Task MergePdfs_ShouldMergeMultiplePdfs()
    {
        // Arrange
        var pdf1Path = CreateTestPdf("test_merge1.pdf");
        var pdf2Path = CreateTestPdf("test_merge2.pdf");
        var outputPath = CreateTestFilePath("test_merge_output.pdf");
        var arguments = new JsonObject
        {
            ["operation"] = "merge",
            ["inputPaths"] = new JsonArray { pdf1Path, pdf2Path },
            ["outputPath"] = outputPath
        };

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.True(File.Exists(outputPath), "Merged PDF should be created");
        var document = new Document(outputPath);
        Assert.True(document.Pages.Count >= 2, "Merged PDF should have multiple pages");
    }

    [Fact]
    public async Task SplitPdf_ShouldSplitIntoMultipleFiles()
    {
        // Arrange
        var pdfPath = CreateTestPdf("test_split.pdf");
        var document = new Document(pdfPath);
        document.Pages.Add();
        document.Save(pdfPath);

        var outputDir = Path.Combine(TestDir, "split_output");
        Directory.CreateDirectory(outputDir);
        var arguments = new JsonObject
        {
            ["operation"] = "split",
            ["path"] = pdfPath,
            ["outputDir"] = outputDir,
            ["pagesPerFile"] = 1
        };

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        var files = Directory.GetFiles(outputDir);
        Assert.True(files.Length >= 2, "Should create multiple files for split pages");
    }

    [Fact]
    public async Task CompressPdf_ShouldCompressPdf()
    {
        // Arrange
        var pdfPath = CreateTestPdf("test_compress.pdf");
        var outputPath = CreateTestFilePath("test_compress_output.pdf");
        var arguments = new JsonObject
        {
            ["operation"] = "compress",
            ["path"] = pdfPath,
            ["outputPath"] = outputPath,
            ["compressImages"] = true
        };

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.True(File.Exists(outputPath), "Compressed PDF should be created");
    }

    [Fact]
    public async Task EncryptPdf_ShouldEncryptPdf()
    {
        // Arrange
        var pdfPath = CreateTestPdf("test_encrypt.pdf");
        var outputPath = CreateTestFilePath("test_encrypt_output.pdf");
        var arguments = new JsonObject
        {
            ["operation"] = "encrypt",
            ["path"] = pdfPath,
            ["outputPath"] = outputPath,
            ["userPassword"] = "user123",
            ["ownerPassword"] = "owner123"
        };

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.True(File.Exists(outputPath), "Encrypted PDF should be created");
    }

    [Fact]
    public async Task CompressPdf_WithCompressFonts_ShouldCompressWithFontSubsetting()
    {
        // Arrange
        var pdfPath = CreateTestPdf("test_compress_fonts.pdf");
        var outputPath = CreateTestFilePath("test_compress_fonts_output.pdf");
        var arguments = new JsonObject
        {
            ["operation"] = "compress",
            ["path"] = pdfPath,
            ["outputPath"] = outputPath,
            ["compressImages"] = true,
            ["compressFonts"] = true
        };

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.True(File.Exists(outputPath), "Compressed PDF with font subsetting should be created");
    }

    [Fact]
    public async Task CompressPdf_WithRemoveUnusedObjects_ShouldRemoveUnused()
    {
        // Arrange
        var pdfPath = CreateTestPdf("test_compress_unused.pdf");
        var outputPath = CreateTestFilePath("test_compress_unused_output.pdf");
        var arguments = new JsonObject
        {
            ["operation"] = "compress",
            ["path"] = pdfPath,
            ["outputPath"] = outputPath,
            ["removeUnusedObjects"] = true
        };

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.True(File.Exists(outputPath), "Compressed PDF with unused objects removed should be created");
    }

    [Fact]
    public async Task CompressPdf_WithAllOptions_ShouldApplyAllCompression()
    {
        // Arrange
        var pdfPath = CreateTestPdf("test_compress_all.pdf");
        var outputPath = CreateTestFilePath("test_compress_all_output.pdf");
        var arguments = new JsonObject
        {
            ["operation"] = "compress",
            ["path"] = pdfPath,
            ["outputPath"] = outputPath,
            ["compressImages"] = true,
            ["compressFonts"] = true,
            ["removeUnusedObjects"] = true
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.True(File.Exists(outputPath), "Fully compressed PDF should be created");
        Assert.Contains("compressed", result, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public async Task CompressPdf_WithNoCompression_ShouldStillCreateOutput()
    {
        // Arrange
        var pdfPath = CreateTestPdf("test_compress_none.pdf");
        var outputPath = CreateTestFilePath("test_compress_none_output.pdf");
        var arguments = new JsonObject
        {
            ["operation"] = "compress",
            ["path"] = pdfPath,
            ["outputPath"] = outputPath,
            ["compressImages"] = false,
            ["compressFonts"] = false,
            ["removeUnusedObjects"] = false
        };

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.True(File.Exists(outputPath), "PDF should be created even with no compression");
    }

    [Fact]
    public async Task SplitPdf_WithMultiplePagesPerFile_ShouldSplitCorrectly()
    {
        // Arrange
        var pdfPath = CreateTestPdf("test_split_multi.pdf");

        var isEvaluationMode = IsEvaluationMode();
        if (isEvaluationMode)
        {
            // In evaluation mode, Aspose.PDF limits collections to 4 elements
            // This test requires more pages than allowed, so skip in evaluation mode
            Assert.True(true, "Test skipped in evaluation mode due to page limit");
            return;
        }

        var document = new Document(pdfPath);
        // Add multiple pages
        document.Pages.Add();
        document.Pages.Add();
        document.Pages.Add();
        document.Pages.Add(); // Total 5 pages
        document.Save(pdfPath);

        var outputDir = Path.Combine(TestDir, "split_multi_output");
        Directory.CreateDirectory(outputDir);
        var arguments = new JsonObject
        {
            ["operation"] = "split",
            ["path"] = pdfPath,
            ["outputDir"] = outputDir,
            ["pagesPerFile"] = 2
        };

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        var files = Directory.GetFiles(outputDir, "*.pdf");
        Assert.True(files.Length >= 2, "Should create multiple files when splitting with 2 pages per file");
    }

    [Fact]
    public async Task MergePdfs_WithThreePdfs_ShouldMergeAll()
    {
        // Arrange
        var pdf1Path = CreateTestPdf("test_merge3_1.pdf");
        var pdf2Path = CreateTestPdf("test_merge3_2.pdf");
        var pdf3Path = CreateTestPdf("test_merge3_3.pdf");
        var outputPath = CreateTestFilePath("test_merge3_output.pdf");
        var arguments = new JsonObject
        {
            ["operation"] = "merge",
            ["inputPaths"] = new JsonArray { pdf1Path, pdf2Path, pdf3Path },
            ["outputPath"] = outputPath
        };

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.True(File.Exists(outputPath), "Merged PDF should be created");
        var document = new Document(outputPath);
        Assert.True(document.Pages.Count >= 3, "Merged PDF should have at least 3 pages");
    }
}