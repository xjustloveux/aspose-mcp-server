using System.Text.Json.Nodes;
using Aspose.Pdf;
using AsposeMcpServer.Tests.Helpers;
using AsposeMcpServer.Tools.Pdf;

namespace AsposeMcpServer.Tests.Pdf;

public class PdfAttachmentToolTests : PdfTestBase
{
    private readonly PdfAttachmentTool _tool = new();

    private string CreateTestPdf(string fileName)
    {
        var filePath = CreateTestFilePath(fileName);
        using var document = new Document();
        document.Pages.Add();
        document.Save(filePath);
        return filePath;
    }

    private string CreateTestAttachment(string fileName)
    {
        var filePath = CreateTestFilePath(fileName);
        File.WriteAllText(filePath, "Test attachment content");
        return filePath;
    }

    [Fact]
    public async Task AddAttachment_ShouldAddAttachment()
    {
        // Arrange
        var pdfPath = CreateTestPdf("test_add_attachment.pdf");
        var attachmentPath = CreateTestAttachment("test_attachment.txt");
        var outputPath = CreateTestFilePath("test_add_attachment_output.pdf");
        var arguments = new JsonObject
        {
            ["operation"] = "add",
            ["path"] = pdfPath,
            ["outputPath"] = outputPath,
            ["attachmentPath"] = attachmentPath,
            ["attachmentName"] = "test_attachment.txt"
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.Contains("Added attachment", result);
        using var document = new Document(outputPath);
        Assert.True(document.EmbeddedFiles.Count > 0, "PDF should contain at least one attachment");
    }

    [Fact]
    public async Task AddAttachment_WithDescription_ShouldAddAttachmentWithDescription()
    {
        // Arrange
        var pdfPath = CreateTestPdf("test_add_attachment_desc.pdf");
        var attachmentPath = CreateTestAttachment("test_attachment_desc.txt");
        var outputPath = CreateTestFilePath("test_add_attachment_desc_output.pdf");
        var arguments = new JsonObject
        {
            ["operation"] = "add",
            ["path"] = pdfPath,
            ["outputPath"] = outputPath,
            ["attachmentPath"] = attachmentPath,
            ["attachmentName"] = "test_attachment_desc.txt",
            ["description"] = "Test description"
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.Contains("Added attachment", result);
    }

    [Fact]
    public async Task AddAttachment_DuplicateName_ShouldThrow()
    {
        // Arrange
        var pdfPath = CreateTestPdf("test_add_duplicate.pdf");
        var attachmentPath = CreateTestAttachment("test_dup.txt");

        using (var document = new Document(pdfPath))
        {
            var fileSpec = new FileSpecification(attachmentPath, "") { Name = "existing.txt" };
            document.EmbeddedFiles.Add(fileSpec);
            document.Save(pdfPath);
        }

        var arguments = new JsonObject
        {
            ["operation"] = "add",
            ["path"] = pdfPath,
            ["attachmentPath"] = attachmentPath,
            ["attachmentName"] = "existing.txt"
        };

        // Act & Assert
        await Assert.ThrowsAsync<ArgumentException>(() => _tool.ExecuteAsync(arguments));
    }

    [Fact]
    public async Task AddAttachment_FileNotFound_ShouldThrow()
    {
        // Arrange
        var pdfPath = CreateTestPdf("test_add_notfound.pdf");
        var arguments = new JsonObject
        {
            ["operation"] = "add",
            ["path"] = pdfPath,
            ["attachmentPath"] = "nonexistent_file.txt",
            ["attachmentName"] = "test.txt"
        };

        // Act & Assert
        await Assert.ThrowsAsync<FileNotFoundException>(() => _tool.ExecuteAsync(arguments));
    }

    [Fact]
    public async Task GetAttachments_WithAttachments_ShouldReturnAttachmentInfo()
    {
        // Arrange
        var pdfPath = CreateTestPdf("test_get_attachments.pdf");
        using (var document = new Document(pdfPath))
        {
            var attachmentPath = CreateTestAttachment("test_attachment2.txt");
            document.EmbeddedFiles.Add(new FileSpecification(attachmentPath, "test_attachment2.txt"));
            document.Save(pdfPath);
        }

        var arguments = new JsonObject
        {
            ["operation"] = "get",
            ["path"] = pdfPath
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.Contains("\"count\": 1", result);
        Assert.Contains("test_attachment2.txt", result);
    }

    [Fact]
    public async Task GetAttachments_Empty_ShouldReturnEmptyResult()
    {
        // Arrange
        var pdfPath = CreateTestPdf("test_get_empty.pdf");
        var arguments = new JsonObject
        {
            ["operation"] = "get",
            ["path"] = pdfPath
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.Contains("\"count\": 0", result);
        Assert.Contains("No attachments found", result);
    }

    [Fact]
    public async Task DeleteAttachment_ShouldDeleteAttachment()
    {
        // Arrange
        var pdfPath = CreateTestPdf("test_delete_attachment.pdf");
        using (var document = new Document(pdfPath))
        {
            var attachmentPath = CreateTestAttachment("test_attachment3.txt");
            document.EmbeddedFiles.Add(new FileSpecification(attachmentPath, "test_attachment3.txt"));
            document.Save(pdfPath);
        }

        var outputPath = CreateTestFilePath("test_delete_attachment_output.pdf");
        var arguments = new JsonObject
        {
            ["operation"] = "delete",
            ["path"] = pdfPath,
            ["outputPath"] = outputPath,
            ["attachmentName"] = "test_attachment3.txt"
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.Contains("Deleted attachment", result);
        using var resultDocument = new Document(outputPath);
        Assert.Empty(resultDocument.EmbeddedFiles);
    }

    [Fact]
    public async Task DeleteAttachment_NotFound_ShouldThrow()
    {
        // Arrange
        var pdfPath = CreateTestPdf("test_delete_notfound.pdf");
        var arguments = new JsonObject
        {
            ["operation"] = "delete",
            ["path"] = pdfPath,
            ["attachmentName"] = "nonexistent.txt"
        };

        // Act & Assert
        var ex = await Assert.ThrowsAsync<ArgumentException>(() => _tool.ExecuteAsync(arguments));
        Assert.Contains("not found", ex.Message);
    }

    [Fact]
    public async Task UnknownOperation_ShouldThrow()
    {
        // Arrange
        var pdfPath = CreateTestPdf("test_unknown_op.pdf");
        var arguments = new JsonObject
        {
            ["operation"] = "unknown",
            ["path"] = pdfPath
        };

        // Act & Assert
        await Assert.ThrowsAsync<ArgumentException>(() => _tool.ExecuteAsync(arguments));
    }
}