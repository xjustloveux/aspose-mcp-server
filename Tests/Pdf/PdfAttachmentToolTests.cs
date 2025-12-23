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
        var document = new Document();
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
        await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.True(File.Exists(outputPath), "Output PDF should be created");
    }

    [Fact]
    public async Task GetAttachments_ShouldReturnAllAttachments()
    {
        // Arrange
        var pdfPath = CreateTestPdf("test_get_attachments.pdf");
        var document = new Document(pdfPath);
        var attachmentPath = CreateTestAttachment("test_attachment2.txt");
        document.EmbeddedFiles.Add(new FileSpecification(attachmentPath));
        document.Save(pdfPath);

        var arguments = new JsonObject
        {
            ["operation"] = "get",
            ["path"] = pdfPath
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.NotNull(result);
        Assert.NotEmpty(result);
        Assert.Contains("Attachment", result, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public async Task DeleteAttachment_ShouldDeleteAttachment()
    {
        // Arrange
        var pdfPath = CreateTestPdf("test_delete_attachment.pdf");
        using (var document = new Document(pdfPath))
        {
            var attachmentPath = CreateTestAttachment("test_attachment3.txt");
            var fileSpec = new FileSpecification(attachmentPath, "test_attachment3.txt");
            document.EmbeddedFiles.Add(fileSpec);
            document.Save(pdfPath);
        }

        // Reload to get accurate count
        int attachmentsBefore;
        using (var document = new Document(pdfPath))
        {
            attachmentsBefore = document.EmbeddedFiles.Count;
            Assert.True(attachmentsBefore > 0, "Attachment should exist before deletion");
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
        await _tool.ExecuteAsync(arguments);

        // Assert
        using var resultDocument = new Document(outputPath);
        var attachmentsAfter = resultDocument.EmbeddedFiles.Count;
        Assert.True(attachmentsAfter < attachmentsBefore,
            $"Attachment should be deleted. Before: {attachmentsBefore}, After: {attachmentsAfter}");
    }
}