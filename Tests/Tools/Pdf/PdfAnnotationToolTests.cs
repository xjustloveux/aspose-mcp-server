using Aspose.Pdf;
using Aspose.Pdf.Annotations;
using AsposeMcpServer.Tests.Helpers;
using AsposeMcpServer.Tools.Pdf;

namespace AsposeMcpServer.Tests.Tools.Pdf;

/// <summary>
///     Integration tests for PdfAnnotationTool.
///     Focuses on session management, file I/O, and operation routing.
///     Detailed parameter validation and business logic tests are in Handler tests.
/// </summary>
public class PdfAnnotationToolTests : PdfTestBase
{
    private readonly PdfAnnotationTool _tool;

    public PdfAnnotationToolTests()
    {
        _tool = new PdfAnnotationTool(SessionManager);
    }

    private string CreateTestPdf(string fileName)
    {
        var filePath = CreateTestFilePath(fileName);
        using var document = new Document();
        document.Pages.Add();
        document.Save(filePath);
        return filePath;
    }

    private string CreatePdfWithAnnotation(string fileName, string content = "Test Note")
    {
        var filePath = CreateTestFilePath(fileName);
        using var document = new Document();
        var page = document.Pages.Add();
        var annotation = new TextAnnotation(page, new Rectangle(100, 100, 200, 130))
        {
            Title = "Test",
            Contents = content
        };
        page.Annotations.Add(annotation);
        document.Save(filePath);
        return filePath;
    }

    private string CreatePdfWithMultipleAnnotations(string fileName, int count = 2)
    {
        var filePath = CreateTestFilePath(fileName);
        using var document = new Document();
        var page = document.Pages.Add();
        for (var i = 0; i < count; i++)
        {
            var annotation = new TextAnnotation(page, new Rectangle(100 + i * 50, 100, 200 + i * 50, 130))
            {
                Contents = $"Note {i + 1}"
            };
            page.Annotations.Add(annotation);
        }

        document.Save(filePath);
        return filePath;
    }

    #region File I/O Smoke Tests

    [Fact]
    public void Add_ShouldAddAnnotation()
    {
        var pdfPath = CreateTestPdf("test_add.pdf");
        var outputPath = CreateTestFilePath("test_add_output.pdf");
        var result = _tool.Execute("add", pdfPath, outputPath: outputPath,
            pageIndex: 1, text: "Test Note", x: 100, y: 100);
        Assert.StartsWith("Added annotation", result);
        using var document = new Document(outputPath);
        Assert.True(document.Pages[1].Annotations.Count > 0);
    }

    [Fact]
    public void Get_WithPageIndex_ShouldReturnPageAnnotations()
    {
        var pdfPath = CreatePdfWithAnnotation("test_get.pdf");
        var result = _tool.Execute("get", pdfPath, pageIndex: 1);
        Assert.Contains("\"type\": \"Text\"", result);
        Assert.Contains("\"pageIndex\": 1", result);
    }

    [Fact]
    public void Delete_WithIndex_ShouldDeleteSingleAnnotation()
    {
        var pdfPath = CreatePdfWithAnnotation("test_delete_single.pdf", "Note to Delete");
        var outputPath = CreateTestFilePath("test_delete_single_output.pdf");
        var result = _tool.Execute("delete", pdfPath, outputPath: outputPath,
            pageIndex: 1, annotationIndex: 1);
        Assert.StartsWith("Deleted annotation 1", result);
        using var document = new Document(outputPath);
        Assert.Empty(document.Pages[1].Annotations);
    }

    [Fact]
    public void Edit_Text_ShouldModifyAnnotation()
    {
        var pdfPath = CreatePdfWithAnnotation("test_edit.pdf", "Original Note");
        var outputPath = CreateTestFilePath("test_edit_output.pdf");
        var result = _tool.Execute("edit", pdfPath, outputPath: outputPath,
            pageIndex: 1, annotationIndex: 1, text: "Updated Note");
        Assert.StartsWith("Annotation 1 on page 1 updated", result);
        using var document = new Document(outputPath);
        var annotation = document.Pages[1].Annotations[1] as TextAnnotation;
        Assert.NotNull(annotation);
        Assert.Equal("Updated Note", annotation.Contents);
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
        var outputPath = CreateTestFilePath($"test_case_{operation}_output.pdf");
        var result = _tool.Execute(operation, pdfPath, outputPath: outputPath,
            pageIndex: 1, text: "Test", x: 100, y: 100);
        Assert.StartsWith("Added annotation", result);
    }

    [Fact]
    public void Execute_WithUnknownOperation_ShouldThrowArgumentException()
    {
        var pdfPath = CreateTestPdf("test_unknown_op.pdf");
        var ex = Assert.Throws<ArgumentException>(() => _tool.Execute("unknown", pdfPath));
        Assert.Contains("Unknown operation", ex.Message);
    }

    [Fact]
    public void Execute_WithNoPathOrSessionId_ShouldThrowException()
    {
        Assert.ThrowsAny<Exception>(() => _tool.Execute("get", pageIndex: 1));
    }

    #endregion

    #region Session Management

    [Fact]
    public void Get_WithSessionId_ShouldGetFromSession()
    {
        var pdfPath = CreatePdfWithAnnotation("test_session_get.pdf", "Session Note");
        var sessionId = OpenSession(pdfPath);
        var result = _tool.Execute("get", sessionId: sessionId, pageIndex: 1);
        Assert.Contains("\"type\": \"Text\"", result);
    }

    [Fact]
    public void Add_WithSessionId_ShouldAddToSession()
    {
        var pdfPath = CreateTestPdf("test_session_add.pdf");
        var sessionId = OpenSession(pdfPath);
        var result = _tool.Execute("add", sessionId: sessionId,
            pageIndex: 1, text: "Session Annotation", x: 150, y: 150);
        Assert.StartsWith("Added annotation", result);
        var document = SessionManager.GetDocument<Document>(sessionId);
        Assert.True(document.Pages[1].Annotations.Count > 0);
    }

    [Fact]
    public void Delete_WithSessionId_ShouldDeleteFromSession()
    {
        var pdfPath = CreatePdfWithMultipleAnnotations("test_session_delete.pdf");
        var sessionId = OpenSession(pdfPath);
        var result = _tool.Execute("delete", sessionId: sessionId, pageIndex: 1);
        Assert.StartsWith("Deleted all 2 annotation(s)", result);
    }

    [Fact]
    public void Edit_WithSessionId_ShouldEditInSession()
    {
        var pdfPath = CreatePdfWithAnnotation("test_session_edit.pdf", "Original");
        var sessionId = OpenSession(pdfPath);
        var result = _tool.Execute("edit", sessionId: sessionId,
            pageIndex: 1, annotationIndex: 1, text: "Edited");
        Assert.StartsWith("Annotation 1 on page 1 updated", result);
        var document = SessionManager.GetDocument<Document>(sessionId);
        var annotation = document.Pages[1].Annotations[1] as TextAnnotation;
        Assert.NotNull(annotation);
        Assert.Equal("Edited", annotation.Contents);
    }

    [Fact]
    public void Execute_WithInvalidSessionId_ShouldThrowKeyNotFoundException()
    {
        Assert.Throws<KeyNotFoundException>(() =>
            _tool.Execute("get", sessionId: "invalid_session", pageIndex: 1));
    }

    [Fact]
    public void Execute_WithBothPathAndSessionId_ShouldPreferSessionId()
    {
        var pdfPath1 = CreateTestPdf("test_path_file.pdf");
        var pdfPath2 = CreatePdfWithAnnotation("test_session_file.pdf", "Session Annotation");
        var sessionId = OpenSession(pdfPath2);
        var result = _tool.Execute("get", pdfPath1, sessionId, pageIndex: 1);
        Assert.Contains("Session Annotation", result);
    }

    #endregion
}
