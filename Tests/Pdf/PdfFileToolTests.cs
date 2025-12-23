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
}