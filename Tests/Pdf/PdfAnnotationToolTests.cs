using System.Text.Json.Nodes;
using Aspose.Pdf;
using Aspose.Pdf.Annotations;
using AsposeMcpServer.Tests.Helpers;
using AsposeMcpServer.Tools.Pdf;

namespace AsposeMcpServer.Tests.Pdf;

public class PdfAnnotationToolTests : PdfTestBase
{
    private readonly PdfAnnotationTool _tool = new();

    private string CreateTestPdf(string fileName)
    {
        var filePath = CreateTestFilePath(fileName);
        var document = new Document();
        document.Pages.Add();
        document.Save(filePath);
        return filePath;
    }

    [Fact]
    public async Task AddAnnotation_ShouldAddAnnotation()
    {
        // Arrange
        var pdfPath = CreateTestPdf("test_add_annotation.pdf");
        var outputPath = CreateTestFilePath("test_add_annotation_output.pdf");
        var arguments = new JsonObject
        {
            ["operation"] = "add",
            ["path"] = pdfPath,
            ["outputPath"] = outputPath,
            ["pageIndex"] = 1,
            ["text"] = "Test Note",
            ["x"] = 100,
            ["y"] = 100
        };

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        var document = new Document(outputPath);
        var page = document.Pages[1];
        var annotations = page.Annotations;
        Assert.True(annotations.Count > 0, "Page should contain at least one annotation");
    }

    [Fact]
    public async Task GetAnnotations_ShouldReturnAllAnnotations()
    {
        // Arrange
        var pdfPath = CreateTestPdf("test_get_annotations.pdf");
        var document = new Document(pdfPath);
        var page = document.Pages[1];
        var annotation = new TextAnnotation(page, new Rectangle(100, 100, 200, 130));
        annotation.Title = "Test";
        annotation.Contents = "Test Note";
        page.Annotations.Add(annotation);
        document.Save(pdfPath);

        var arguments = new JsonObject
        {
            ["operation"] = "get",
            ["path"] = pdfPath,
            ["pageIndex"] = 1
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.NotNull(result);
        Assert.NotEmpty(result);
        Assert.Contains("Annotation", result, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public async Task DeleteAnnotation_ShouldDeleteAnnotation()
    {
        // Arrange
        var pdfPath = CreateTestPdf("test_delete_annotation.pdf");
        var document = new Document(pdfPath);
        var page = document.Pages[1];
        var annotation = new TextAnnotation(page, new Rectangle(100, 100, 200, 130));
        annotation.Contents = "Note to Delete";
        page.Annotations.Add(annotation);
        document.Save(pdfPath);

        var annotationsBefore = document.Pages[1].Annotations.Count;
        Assert.True(annotationsBefore > 0, "Annotation should exist before deletion");

        var outputPath = CreateTestFilePath("test_delete_annotation_output.pdf");
        var arguments = new JsonObject
        {
            ["operation"] = "delete",
            ["path"] = pdfPath,
            ["outputPath"] = outputPath,
            ["pageIndex"] = 1,
            ["annotationIndex"] = 1
        };

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        var resultDocument = new Document(outputPath);
        var annotationsAfter = resultDocument.Pages[1].Annotations.Count;
        Assert.True(annotationsAfter < annotationsBefore,
            $"Annotation should be deleted. Before: {annotationsBefore}, After: {annotationsAfter}");
    }

    [Fact]
    public async Task EditAnnotation_ShouldModifyAnnotation()
    {
        // Arrange
        var pdfPath = CreateTestPdf("test_edit_annotation.pdf");
        var document = new Document(pdfPath);
        var page = document.Pages[1];
        var annotation = new TextAnnotation(page, new Rectangle(100, 100, 200, 130));
        annotation.Contents = "Original Note";
        page.Annotations.Add(annotation);
        document.Save(pdfPath);

        var outputPath = CreateTestFilePath("test_edit_annotation_output.pdf");
        var arguments = new JsonObject
        {
            ["operation"] = "edit",
            ["path"] = pdfPath,
            ["outputPath"] = outputPath,
            ["pageIndex"] = 1,
            ["annotationIndex"] = 1,
            ["text"] = "Updated Note"
        };

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.True(File.Exists(outputPath), "Output PDF should be created");
        var resultDocument = new Document(outputPath);
        var resultPage = resultDocument.Pages[1];
        Assert.True(resultPage.Annotations.Count > 0, "Page should contain annotations");
        var editedAnnotation = resultPage.Annotations[1] as TextAnnotation;
        Assert.NotNull(editedAnnotation);
        Assert.Contains("Updated", editedAnnotation.Contents, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("Original", editedAnnotation.Contents, StringComparison.OrdinalIgnoreCase);
    }
}