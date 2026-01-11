using Aspose.Pdf;
using AsposeMcpServer.Tests.Helpers;
using AsposeMcpServer.Tools.Pdf;

namespace AsposeMcpServer.Tests.Tools.Pdf;

/// <summary>
///     Integration tests for PdfAttachmentTool.
///     Focuses on session management, file I/O, and operation routing.
///     Detailed parameter validation and business logic tests are in Handler tests.
/// </summary>
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

    #region File I/O Smoke Tests

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
    public void Get_WithAttachments_ShouldReturnAttachmentInfo()
    {
        var pdfPath = CreatePdfWithAttachment("test_get.pdf", "test_attachment.txt");
        var result = _tool.Execute("get", pdfPath);
        Assert.Contains("\"count\": 1", result);
        Assert.Contains("test_attachment.txt", result);
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

    #endregion

    #region Operation Routing

    [Theory]
    [InlineData("ADD")]
    [InlineData("Add")]
    [InlineData("add")]
    public void Operation_ShouldBeCaseInsensitive(string operation)
    {
        var pdfPath = CreateTestPdf($"test_case_{operation}.pdf");
        var attachmentPath = CreateTestAttachment($"test_case_{operation}.txt");
        var outputPath = CreateTestFilePath($"test_case_{operation}_output.pdf");
        var result = _tool.Execute(operation, pdfPath, outputPath: outputPath,
            attachmentPath: attachmentPath, attachmentName: $"attachment_{operation}.txt");
        Assert.StartsWith("Added attachment", result);
    }

    [Fact]
    public void Execute_WithUnknownOperation_ShouldThrowArgumentException()
    {
        var pdfPath = CreateTestPdf("test_unknown_op.pdf");
        var ex = Assert.Throws<ArgumentException>(() => _tool.Execute("unknown", pdfPath));
        Assert.Contains("Unknown operation", ex.Message);
    }

    #endregion

    #region Session Management

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
    public void Delete_WithSessionId_ShouldDeleteFromSession()
    {
        var pdfPath = CreatePdfWithAttachment("test_session_delete.pdf", "to_delete.txt");
        var sessionId = OpenSession(pdfPath);
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
