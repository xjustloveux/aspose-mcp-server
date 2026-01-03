using Aspose.Pdf;
using AsposeMcpServer.Tests.Helpers;
using AsposeMcpServer.Tools.Pdf;

namespace AsposeMcpServer.Tests.Tools.Pdf;

public class PdfAttachmentToolTests : PdfTestBase
{
    private readonly PdfAttachmentTool _tool;

    public PdfAttachmentToolTests()
    {
        _tool = new PdfAttachmentTool(SessionManager);
    }

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

    #region General Tests

    [Fact]
    public void AddAttachment_ShouldAddAttachment()
    {
        var pdfPath = CreateTestPdf("test_add_attachment.pdf");
        var attachmentPath = CreateTestAttachment("test_attachment.txt");
        var outputPath = CreateTestFilePath("test_add_attachment_output.pdf");
        var result = _tool.Execute(
            "add",
            pdfPath,
            outputPath: outputPath,
            attachmentPath: attachmentPath,
            attachmentName: "test_attachment.txt");
        Assert.Contains("Added attachment", result);
        using var document = new Document(outputPath);
        Assert.True(document.EmbeddedFiles.Count > 0, "PDF should contain at least one attachment");
    }

    [Fact]
    public void AddAttachment_WithDescription_ShouldAddAttachmentWithDescription()
    {
        var pdfPath = CreateTestPdf("test_add_attachment_desc.pdf");
        var attachmentPath = CreateTestAttachment("test_attachment_desc.txt");
        var outputPath = CreateTestFilePath("test_add_attachment_desc_output.pdf");
        var result = _tool.Execute(
            "add",
            pdfPath,
            outputPath: outputPath,
            attachmentPath: attachmentPath,
            attachmentName: "test_attachment_desc.txt",
            description: "Test description");
        Assert.Contains("Added attachment", result);
    }

    [Fact]
    public void AddAttachment_DuplicateName_ShouldThrow()
    {
        var pdfPath = CreateTestPdf("test_add_duplicate.pdf");
        var attachmentPath = CreateTestAttachment("test_dup.txt");

        using (var document = new Document(pdfPath))
        {
            var fileSpec = new FileSpecification(attachmentPath, "") { Name = "existing.txt" };
            document.EmbeddedFiles.Add(fileSpec);
            document.Save(pdfPath);
        }

        Assert.Throws<ArgumentException>(() => _tool.Execute(
            "add",
            pdfPath,
            attachmentPath: attachmentPath,
            attachmentName: "existing.txt"));
    }

    [Fact]
    public void AddAttachment_FileNotFound_ShouldThrow()
    {
        var pdfPath = CreateTestPdf("test_add_notfound.pdf");
        Assert.Throws<FileNotFoundException>(() => _tool.Execute(
            "add",
            pdfPath,
            attachmentPath: "nonexistent_file.txt",
            attachmentName: "test.txt"));
    }

    [Fact]
    public void GetAttachments_WithAttachments_ShouldReturnAttachmentInfo()
    {
        var pdfPath = CreateTestPdf("test_get_attachments.pdf");
        using (var document = new Document(pdfPath))
        {
            var attachmentPath = CreateTestAttachment("test_attachment2.txt");
            document.EmbeddedFiles.Add(new FileSpecification(attachmentPath, "test_attachment2.txt"));
            document.Save(pdfPath);
        }

        var result = _tool.Execute("get", pdfPath);
        Assert.Contains("\"count\": 1", result);
        Assert.Contains("test_attachment2.txt", result);
    }

    [Fact]
    public void GetAttachments_Empty_ShouldReturnEmptyResult()
    {
        var pdfPath = CreateTestPdf("test_get_empty.pdf");
        var result = _tool.Execute("get", pdfPath);
        Assert.Contains("\"count\": 0", result);
        Assert.Contains("No attachments found", result);
    }

    [Fact]
    public void DeleteAttachment_ShouldDeleteAttachment()
    {
        var pdfPath = CreateTestPdf("test_delete_attachment.pdf");
        using (var document = new Document(pdfPath))
        {
            var attachmentPath = CreateTestAttachment("test_attachment3.txt");
            document.EmbeddedFiles.Add(new FileSpecification(attachmentPath, "test_attachment3.txt"));
            document.Save(pdfPath);
        }

        var outputPath = CreateTestFilePath("test_delete_attachment_output.pdf");
        var result = _tool.Execute(
            "delete",
            pdfPath,
            outputPath: outputPath,
            attachmentName: "test_attachment3.txt");
        Assert.Contains("Deleted attachment", result);
        using var resultDocument = new Document(outputPath);
        Assert.Empty(resultDocument.EmbeddedFiles);
    }

    [Fact]
    public void DeleteAttachment_NotFound_ShouldThrow()
    {
        var pdfPath = CreateTestPdf("test_delete_notfound.pdf");
        var ex = Assert.Throws<ArgumentException>(() => _tool.Execute(
            "delete",
            pdfPath,
            attachmentName: "nonexistent.txt"));
        Assert.Contains("not found", ex.Message);
    }

    [Fact]
    public void UnknownOperation_ShouldThrow()
    {
        var pdfPath = CreateTestPdf("test_unknown_op.pdf");
        Assert.Throws<ArgumentException>(() => _tool.Execute("unknown", pdfPath));
    }

    #endregion

    #region Exception Tests

    [Fact]
    public void Execute_WithUnknownOperation_ShouldThrowArgumentException()
    {
        var pdfPath = CreateTestPdf("test_exception_unknown.pdf");
        var exception = Assert.Throws<ArgumentException>(() => _tool.Execute("invalid_operation", pdfPath));
        Assert.Contains("Unknown operation", exception.Message);
    }

    [Fact]
    public void AddAttachment_WithMissingAttachmentPath_ShouldThrow()
    {
        var pdfPath = CreateTestPdf("test_missing_attachment_path.pdf");

        // Act & Assert - missing attachmentPath
        Assert.Throws<ArgumentException>(() => _tool.Execute(
            "add",
            pdfPath,
            attachmentName: "test.txt"));
    }

    #endregion

    #region Session ID Tests

    [Fact]
    public void GetAttachments_WithSessionId_ShouldGetFromSession()
    {
        var pdfPath = CreateTestPdf("test_session_get.pdf");
        using (var document = new Document(pdfPath))
        {
            var attachmentPath = CreateTestAttachment("session_attachment.txt");
            document.EmbeddedFiles.Add(new FileSpecification(attachmentPath, "session_attachment.txt"));
            document.Save(pdfPath);
        }

        var sessionId = OpenSession(pdfPath);
        var result = _tool.Execute("get", sessionId: sessionId);
        Assert.NotNull(result);
        Assert.Contains("session_attachment.txt", result);
    }

    [Fact]
    public void AddAttachment_WithSessionId_ShouldAddToSession()
    {
        var pdfPath = CreateTestPdf("test_session_add.pdf");
        var attachmentPath = CreateTestAttachment("new_session_attachment.txt");
        var sessionId = OpenSession(pdfPath);
        var result = _tool.Execute(
            "add",
            sessionId: sessionId,
            attachmentPath: attachmentPath,
            attachmentName: "new_session_attachment.txt");
        Assert.Contains("Added attachment", result);
        Assert.Contains("session", result);
    }

    [Fact]
    public void AddAttachment_WithSessionId_ShouldModifyInMemory()
    {
        var pdfPath = CreateTestPdf("test_session_memory.pdf");
        var attachmentPath = CreateTestAttachment("memory_attachment.txt");
        var sessionId = OpenSession(pdfPath);
        _tool.Execute(
            "add",
            sessionId: sessionId,
            attachmentPath: attachmentPath,
            attachmentName: "memory_attachment.txt");

        // Assert - verify in-memory changes
        var document = SessionManager.GetDocument<Document>(sessionId);
        Assert.NotNull(document);
        Assert.True(document.EmbeddedFiles.Count > 0);
    }

    #endregion
}