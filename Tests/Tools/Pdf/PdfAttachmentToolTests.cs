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

    private string CreatePdfWithAttachment(string fileName, string attachmentName = "existing.txt")
    {
        var filePath = CreateTestFilePath(fileName);
        using var document = new Document();
        document.Pages.Add();
        var attachmentPath = CreateTestAttachment(attachmentName);
        document.EmbeddedFiles.Add(new FileSpecification(attachmentPath, attachmentName));
        document.Save(filePath);
        return filePath;
    }

    #region General

    [Fact]
    public void Add_ShouldAddAttachment()
    {
        var pdfPath = CreateTestPdf("test_add.pdf");
        var attachmentPath = CreateTestAttachment("test_attachment.txt");
        var outputPath = CreateTestFilePath("test_add_output.pdf");
        var result = _tool.Execute("add", pdfPath, outputPath: outputPath,
            attachmentPath: attachmentPath, attachmentName: "test_attachment.txt");
        Assert.StartsWith("Added attachment", result);
        using var document = new Document(outputPath);
        Assert.True(document.EmbeddedFiles.Count > 0);
    }

    [Fact]
    public void Add_WithDescription_ShouldAddAttachmentWithDescription()
    {
        var pdfPath = CreateTestPdf("test_add_desc.pdf");
        var attachmentPath = CreateTestAttachment("test_attachment_desc.txt");
        var outputPath = CreateTestFilePath("test_add_desc_output.pdf");
        var result = _tool.Execute("add", pdfPath, outputPath: outputPath,
            attachmentPath: attachmentPath, attachmentName: "test_attachment_desc.txt",
            description: "Test description");
        Assert.StartsWith("Added attachment", result);
        Assert.True(File.Exists(outputPath));
    }

    [Fact]
    public void Get_WithAttachments_ShouldReturnAttachmentInfo()
    {
        var pdfPath = CreatePdfWithAttachment("test_get.pdf", "test_attachment.txt");
        var result = _tool.Execute("get", pdfPath);
        Assert.Contains("\"count\": 1", result);
        Assert.Contains("test_attachment.txt", result);
    }

    [Fact]
    public void Get_WithNoAttachments_ShouldReturnEmptyResult()
    {
        var pdfPath = CreateTestPdf("test_get_empty.pdf");
        var result = _tool.Execute("get", pdfPath);
        Assert.Contains("\"count\": 0", result);
        Assert.Contains("No attachments found", result);
    }

    [Fact]
    public void Delete_ShouldDeleteAttachment()
    {
        var pdfPath = CreatePdfWithAttachment("test_delete.pdf", "to_delete.txt");
        var outputPath = CreateTestFilePath("test_delete_output.pdf");
        var result = _tool.Execute("delete", pdfPath, outputPath: outputPath,
            attachmentName: "to_delete.txt");
        Assert.StartsWith("Deleted attachment", result);
        using var document = new Document(outputPath);
        Assert.Empty(document.EmbeddedFiles);
    }

    [Theory]
    [InlineData("ADD")]
    [InlineData("Add")]
    [InlineData("add")]
    public void Operation_ShouldBeCaseInsensitive_Add(string operation)
    {
        var pdfPath = CreateTestPdf($"test_case_{operation}.pdf");
        var attachmentPath = CreateTestAttachment($"test_case_{operation}.txt");
        var outputPath = CreateTestFilePath($"test_case_{operation}_output.pdf");
        var result = _tool.Execute(operation, pdfPath, outputPath: outputPath,
            attachmentPath: attachmentPath, attachmentName: $"attachment_{operation}.txt");
        Assert.StartsWith("Added attachment", result);
    }

    [Theory]
    [InlineData("GET")]
    [InlineData("Get")]
    [InlineData("get")]
    public void Operation_ShouldBeCaseInsensitive_Get(string operation)
    {
        var pdfPath = CreateTestPdf($"test_case_get_{operation}.pdf");
        var result = _tool.Execute(operation, pdfPath);
        Assert.Contains("count", result);
    }

    #endregion

    #region Exception

    [Fact]
    public void Execute_WithUnknownOperation_ShouldThrowArgumentException()
    {
        var pdfPath = CreateTestPdf("test_unknown_op.pdf");
        var ex = Assert.Throws<ArgumentException>(() => _tool.Execute("unknown", pdfPath));
        Assert.Contains("Unknown operation", ex.Message);
    }

    [Fact]
    public void Add_WithMissingAttachmentPath_ShouldThrowArgumentException()
    {
        var pdfPath = CreateTestPdf("test_add_no_path.pdf");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("add", pdfPath, attachmentName: "test.txt"));
        Assert.Contains("attachmentPath is required", ex.Message);
    }

    [Fact]
    public void Add_WithMissingAttachmentName_ShouldThrowArgumentException()
    {
        var pdfPath = CreateTestPdf("test_add_no_name.pdf");
        var attachmentPath = CreateTestAttachment("test_no_name.txt");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("add", pdfPath, attachmentPath: attachmentPath));
        Assert.Contains("attachmentName is required", ex.Message);
    }

    [Fact]
    public void Add_WithNonExistentFile_ShouldThrowFileNotFoundException()
    {
        var pdfPath = CreateTestPdf("test_add_notfound.pdf");
        Assert.Throws<FileNotFoundException>(() =>
            _tool.Execute("add", pdfPath, attachmentPath: "nonexistent_file.txt",
                attachmentName: "test.txt"));
    }

    [Fact]
    public void Add_WithDuplicateName_ShouldThrowArgumentException()
    {
        var pdfPath = CreatePdfWithAttachment("test_add_duplicate.pdf");
        var attachmentPath = CreateTestAttachment("test_dup.txt");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("add", pdfPath, attachmentPath: attachmentPath,
                attachmentName: "existing.txt"));
        Assert.Contains("already exists", ex.Message);
    }

    [Fact]
    public void Delete_WithMissingAttachmentName_ShouldThrowArgumentException()
    {
        var pdfPath = CreatePdfWithAttachment("test_delete_no_name.pdf");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("delete", pdfPath));
        Assert.Contains("attachmentName is required", ex.Message);
    }

    [Fact]
    public void Delete_WithNonExistentAttachment_ShouldThrowArgumentException()
    {
        var pdfPath = CreateTestPdf("test_delete_notfound.pdf");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("delete", pdfPath, attachmentName: "nonexistent.txt"));
        Assert.Contains("not found", ex.Message);
    }

    [Fact]
    public void Execute_WithNoPathOrSessionId_ShouldThrowException()
    {
        Assert.ThrowsAny<Exception>(() => _tool.Execute("get"));
    }

    #endregion

    #region Session

    [Fact]
    public void Get_WithSessionId_ShouldGetFromSession()
    {
        var pdfPath = CreatePdfWithAttachment("test_session_get.pdf", "session_attachment.txt");
        var sessionId = OpenSession(pdfPath);
        var result = _tool.Execute("get", sessionId: sessionId);
        Assert.Contains("session_attachment.txt", result);
    }

    [Fact]
    public void Add_WithSessionId_ShouldAddToSession()
    {
        var pdfPath = CreateTestPdf("test_session_add.pdf");
        var attachmentPath = CreateTestAttachment("session_attachment.txt");
        var sessionId = OpenSession(pdfPath);
        var result = _tool.Execute("add", sessionId: sessionId,
            attachmentPath: attachmentPath, attachmentName: "session_attachment.txt");
        Assert.StartsWith("Added attachment", result);
        Assert.Contains("session", result);
    }

    [Fact]
    public void Add_WithSessionId_ShouldModifyInMemory()
    {
        var pdfPath = CreateTestPdf("test_session_memory.pdf");
        var attachmentPath = CreateTestAttachment("memory_attachment.txt");
        var sessionId = OpenSession(pdfPath);
        _tool.Execute("add", sessionId: sessionId,
            attachmentPath: attachmentPath, attachmentName: "memory_attachment.txt");
        var document = SessionManager.GetDocument<Document>(sessionId);
        Assert.NotNull(document);
        Assert.True(document.EmbeddedFiles.Count > 0);
    }

    [Fact]
    public void Delete_WithSessionId_ShouldDeleteFromSession()
    {
        var pdfPath = CreatePdfWithAttachment("test_session_delete.pdf", "to_delete.txt");
        var sessionId = OpenSession(pdfPath);
        var docBefore = SessionManager.GetDocument<Document>(sessionId);
        Assert.True(docBefore.EmbeddedFiles.Count > 0);
        var result = _tool.Execute("delete", sessionId: sessionId, attachmentName: "to_delete.txt");
        Assert.StartsWith("Deleted attachment", result);
        Assert.Contains("session", result);
        var docAfter = SessionManager.GetDocument<Document>(sessionId);
        Assert.Empty(docAfter.EmbeddedFiles);
    }

    [Fact]
    public void Execute_WithInvalidSessionId_ShouldThrowKeyNotFoundException()
    {
        Assert.Throws<KeyNotFoundException>(() =>
            _tool.Execute("get", sessionId: "invalid_session"));
    }

    [Fact]
    public void Execute_WithBothPathAndSessionId_ShouldPreferSessionId()
    {
        var pdfPath1 = CreateTestPdf("test_path_file.pdf");
        var pdfPath2 = CreatePdfWithAttachment("test_session_file.pdf", "session_file.txt");
        var sessionId = OpenSession(pdfPath2);
        var result = _tool.Execute("get", pdfPath1, sessionId);
        Assert.Contains("session_file.txt", result);
    }

    #endregion
}