using Aspose.Pdf;
using Aspose.Pdf.Annotations;
using AsposeMcpServer.Tests.Helpers;
using AsposeMcpServer.Tools.Pdf;

namespace AsposeMcpServer.Tests.Tools.Pdf;

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

    #region General Tests

    [Fact]
    public void AddAnnotation_ShouldAddAnnotation()
    {
        var pdfPath = CreateTestPdf("test_add_annotation.pdf");
        var outputPath = CreateTestFilePath("test_add_annotation_output.pdf");
        var result = _tool.Execute(
            "add",
            pdfPath,
            pageIndex: 1,
            outputPath: outputPath,
            text: "Test Note",
            x: 100,
            y: 100);
        Assert.Contains("Added annotation", result);
        using var document = new Document(outputPath);
        var page = document.Pages[1];
        Assert.True(page.Annotations.Count > 0, "Page should contain at least one annotation");
        var annotation = page.Annotations[1] as TextAnnotation;
        Assert.NotNull(annotation);
        Assert.Equal("Test Note", annotation.Contents);
    }

    [Fact]
    public void AddAnnotation_InvalidPageIndex_ShouldThrow()
    {
        var pdfPath = CreateTestPdf("test_add_invalid_page.pdf");
        Assert.Throws<ArgumentException>(() => _tool.Execute(
            "add",
            pdfPath,
            pageIndex: 99,
            text: "Test"));
    }

    [Fact]
    public void GetAnnotations_WithPageIndex_ShouldReturnPageAnnotations()
    {
        var pdfPath = CreateTestPdf("test_get_annotations.pdf");
        using (var document = new Document(pdfPath))
        {
            var page = document.Pages[1];
            var annotation = new TextAnnotation(page, new Rectangle(100, 100, 200, 130))
            {
                Title = "Test",
                Contents = "Test Note"
            };
            page.Annotations.Add(annotation);
            document.Save(pdfPath);
        }

        var result = _tool.Execute("get", pdfPath, pageIndex: 1);
        Assert.NotNull(result);
        Assert.Contains("TextAnnotation", result);
        Assert.Contains("\"pageIndex\": 1", result);
    }

    [Fact]
    public void GetAnnotations_WithoutPageIndex_ShouldReturnAllAnnotations()
    {
        var pdfPath = CreateTestFilePath("test_get_all_annotations.pdf");
        using (var document = new Document())
        {
            document.Pages.Add();
            document.Pages.Add();
            var page1 = document.Pages[1];
            var page2 = document.Pages[2];
            page1.Annotations.Add(new TextAnnotation(page1, new Rectangle(100, 100, 200, 130)) { Contents = "Note 1" });
            page2.Annotations.Add(new TextAnnotation(page2, new Rectangle(100, 100, 200, 130)) { Contents = "Note 2" });
            document.Save(pdfPath);
        }

        var result = _tool.Execute("get", pdfPath);
        Assert.Contains("\"count\": 2", result);
    }

    [Fact]
    public void GetAnnotations_InvalidPageIndex_ShouldThrow()
    {
        var pdfPath = CreateTestPdf("test_get_invalid_page.pdf");
        Assert.Throws<ArgumentException>(() => _tool.Execute("get", pdfPath, pageIndex: 99));
    }

    [Fact]
    public void DeleteAnnotation_WithIndex_ShouldDeleteSingleAnnotation()
    {
        var pdfPath = CreateTestPdf("test_delete_annotation.pdf");
        using (var document = new Document(pdfPath))
        {
            var page = document.Pages[1];
            page.Annotations.Add(new TextAnnotation(page, new Rectangle(100, 100, 200, 130))
                { Contents = "Note to Delete" });
            document.Save(pdfPath);
        }

        var outputPath = CreateTestFilePath("test_delete_annotation_output.pdf");
        var result = _tool.Execute(
            "delete",
            pdfPath,
            pageIndex: 1,
            outputPath: outputPath,
            annotationIndex: 1);
        Assert.Contains("Deleted annotation 1", result);
        using var resultDocument = new Document(outputPath);
        Assert.Empty(resultDocument.Pages[1].Annotations);
    }

    [Fact]
    public void DeleteAnnotation_WithoutIndex_ShouldDeleteAllAnnotations()
    {
        var pdfPath = CreateTestPdf("test_delete_all_annotations.pdf");
        using (var document = new Document(pdfPath))
        {
            var page = document.Pages[1];
            page.Annotations.Add(new TextAnnotation(page, new Rectangle(100, 100, 200, 130)) { Contents = "Note 1" });
            page.Annotations.Add(new TextAnnotation(page, new Rectangle(200, 200, 300, 230)) { Contents = "Note 2" });
            document.Save(pdfPath);
        }

        var outputPath = CreateTestFilePath("test_delete_all_annotations_output.pdf");
        var result = _tool.Execute(
            "delete",
            pdfPath,
            pageIndex: 1,
            outputPath: outputPath);
        Assert.Contains("Deleted all 2 annotation(s)", result);
        using var resultDocument = new Document(outputPath);
        Assert.Empty(resultDocument.Pages[1].Annotations);
    }

    [Fact]
    public void DeleteAnnotation_NoAnnotations_ShouldThrow()
    {
        var pdfPath = CreateTestPdf("test_delete_no_annotations.pdf");
        Assert.Throws<ArgumentException>(() => _tool.Execute("delete", pdfPath, pageIndex: 1));
    }

    [Fact]
    public void DeleteAnnotation_InvalidAnnotationIndex_ShouldThrow()
    {
        var pdfPath = CreateTestPdf("test_delete_invalid_index.pdf");
        using (var document = new Document(pdfPath))
        {
            var page = document.Pages[1];
            page.Annotations.Add(new TextAnnotation(page, new Rectangle(100, 100, 200, 130)) { Contents = "Note" });
            document.Save(pdfPath);
        }

        Assert.Throws<ArgumentException>(() => _tool.Execute(
            "delete",
            pdfPath,
            pageIndex: 1,
            annotationIndex: 99));
    }

    [Fact]
    public void EditAnnotation_ShouldModifyAnnotation()
    {
        var pdfPath = CreateTestPdf("test_edit_annotation.pdf");
        using (var document = new Document(pdfPath))
        {
            var page = document.Pages[1];
            page.Annotations.Add(new TextAnnotation(page, new Rectangle(100, 100, 200, 130))
                { Contents = "Original Note" });
            document.Save(pdfPath);
        }

        var outputPath = CreateTestFilePath("test_edit_annotation_output.pdf");
        var result = _tool.Execute(
            "edit",
            pdfPath,
            pageIndex: 1,
            outputPath: outputPath,
            annotationIndex: 1,
            text: "Updated Note");
        Assert.Contains("Edited annotation", result);
        using var resultDocument = new Document(outputPath);
        var editedAnnotation = resultDocument.Pages[1].Annotations[1] as TextAnnotation;
        Assert.NotNull(editedAnnotation);
        Assert.Equal("Updated Note", editedAnnotation.Contents);
    }

    [Fact]
    public void EditAnnotation_WithPosition_ShouldUpdatePosition()
    {
        var pdfPath = CreateTestPdf("test_edit_position.pdf");
        using (var document = new Document(pdfPath))
        {
            var page = document.Pages[1];
            page.Annotations.Add(new TextAnnotation(page, new Rectangle(100, 100, 300, 150)) { Contents = "Note" });
            document.Save(pdfPath);
        }

        var outputPath = CreateTestFilePath("test_edit_position_output.pdf");
        var result = _tool.Execute(
            "edit",
            pdfPath,
            pageIndex: 1,
            outputPath: outputPath,
            annotationIndex: 1,
            text: "Moved Note",
            x: 200,
            y: 500);
        Assert.Contains("Edited annotation", result);
        using var resultDocument = new Document(outputPath);
        var editedAnnotation = resultDocument.Pages[1].Annotations[1] as TextAnnotation;
        Assert.NotNull(editedAnnotation);
        Assert.Equal(200, editedAnnotation.Rect.LLX, 1);
        Assert.Equal(500, editedAnnotation.Rect.LLY, 1);
    }

    [Fact]
    public void EditAnnotation_InvalidAnnotationIndex_ShouldThrow()
    {
        var pdfPath = CreateTestPdf("test_edit_invalid_index.pdf");
        using (var document = new Document(pdfPath))
        {
            var page = document.Pages[1];
            page.Annotations.Add(new TextAnnotation(page, new Rectangle(100, 100, 200, 130)) { Contents = "Note" });
            document.Save(pdfPath);
        }

        Assert.Throws<ArgumentException>(() => _tool.Execute(
            "edit",
            pdfPath,
            pageIndex: 1,
            annotationIndex: 99,
            text: "Test"));
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
    public void AddAnnotation_WithMissingRequiredParams_ShouldThrow()
    {
        var pdfPath = CreateTestPdf("test_missing_params.pdf");

        // Act & Assert - missing pageIndex
        Assert.Throws<ArgumentException>(() => _tool.Execute(
            "add",
            pdfPath,
            text: "Test Note"));
    }

    #endregion

    #region Session ID Tests

    [Fact]
    public void GetAnnotations_WithSessionId_ShouldGetFromSession()
    {
        var pdfPath = CreateTestPdf("test_session_get.pdf");
        using (var document = new Document(pdfPath))
        {
            var page = document.Pages[1];
            page.Annotations.Add(new TextAnnotation(page, new Rectangle(100, 100, 200, 130))
                { Contents = "Session Note" });
            document.Save(pdfPath);
        }

        var sessionId = OpenSession(pdfPath);
        var result = _tool.Execute("get", sessionId: sessionId, pageIndex: 1);
        Assert.NotNull(result);
        Assert.Contains("TextAnnotation", result);
    }

    [Fact]
    public void AddAnnotation_WithSessionId_ShouldAddToSession()
    {
        var pdfPath = CreateTestPdf("test_session_add.pdf");
        var sessionId = OpenSession(pdfPath);
        var result = _tool.Execute(
            "add",
            sessionId: sessionId,
            pageIndex: 1,
            text: "Session Annotation",
            x: 150,
            y: 150);
        Assert.Contains("Added annotation", result);
        Assert.Contains("session", result);
    }

    [Fact]
    public void AddAnnotation_WithSessionId_ShouldModifyInMemory()
    {
        var pdfPath = CreateTestPdf("test_session_memory.pdf");
        var sessionId = OpenSession(pdfPath);
        _tool.Execute(
            "add",
            sessionId: sessionId,
            pageIndex: 1,
            text: "In-Memory Annotation",
            x: 200,
            y: 200);

        // Assert - verify in-memory changes
        var document = SessionManager.GetDocument<Document>(sessionId);
        Assert.NotNull(document);
        Assert.True(document.Pages[1].Annotations.Count > 0);
    }

    #endregion
}