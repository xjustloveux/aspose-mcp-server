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
        using var document = new Document();
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
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.Contains("Added annotation", result);
        using var document = new Document(outputPath);
        var page = document.Pages[1];
        Assert.True(page.Annotations.Count > 0, "Page should contain at least one annotation");
    }

    [Fact]
    public async Task AddAnnotation_InvalidPageIndex_ShouldThrow()
    {
        // Arrange
        var pdfPath = CreateTestPdf("test_add_invalid_page.pdf");
        var arguments = new JsonObject
        {
            ["operation"] = "add",
            ["path"] = pdfPath,
            ["pageIndex"] = 99,
            ["text"] = "Test"
        };

        // Act & Assert
        await Assert.ThrowsAsync<ArgumentException>(() => _tool.ExecuteAsync(arguments));
    }

    [Fact]
    public async Task GetAnnotations_WithPageIndex_ShouldReturnPageAnnotations()
    {
        // Arrange
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
        Assert.Contains("TextAnnotation", result);
        Assert.Contains("\"pageIndex\": 1", result);
    }

    [Fact]
    public async Task GetAnnotations_WithoutPageIndex_ShouldReturnAllAnnotations()
    {
        // Arrange
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

        var arguments = new JsonObject
        {
            ["operation"] = "get",
            ["path"] = pdfPath
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.Contains("\"count\": 2", result);
    }

    [Fact]
    public async Task GetAnnotations_InvalidPageIndex_ShouldThrow()
    {
        // Arrange
        var pdfPath = CreateTestPdf("test_get_invalid_page.pdf");
        var arguments = new JsonObject
        {
            ["operation"] = "get",
            ["path"] = pdfPath,
            ["pageIndex"] = 99
        };

        // Act & Assert
        await Assert.ThrowsAsync<ArgumentException>(() => _tool.ExecuteAsync(arguments));
    }

    [Fact]
    public async Task DeleteAnnotation_WithIndex_ShouldDeleteSingleAnnotation()
    {
        // Arrange
        var pdfPath = CreateTestPdf("test_delete_annotation.pdf");
        using (var document = new Document(pdfPath))
        {
            var page = document.Pages[1];
            page.Annotations.Add(new TextAnnotation(page, new Rectangle(100, 100, 200, 130))
                { Contents = "Note to Delete" });
            document.Save(pdfPath);
        }

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
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.Contains("Deleted annotation 1", result);
        using var resultDocument = new Document(outputPath);
        Assert.Empty(resultDocument.Pages[1].Annotations);
    }

    [Fact]
    public async Task DeleteAnnotation_WithoutIndex_ShouldDeleteAllAnnotations()
    {
        // Arrange
        var pdfPath = CreateTestPdf("test_delete_all_annotations.pdf");
        using (var document = new Document(pdfPath))
        {
            var page = document.Pages[1];
            page.Annotations.Add(new TextAnnotation(page, new Rectangle(100, 100, 200, 130)) { Contents = "Note 1" });
            page.Annotations.Add(new TextAnnotation(page, new Rectangle(200, 200, 300, 230)) { Contents = "Note 2" });
            document.Save(pdfPath);
        }

        var outputPath = CreateTestFilePath("test_delete_all_annotations_output.pdf");
        var arguments = new JsonObject
        {
            ["operation"] = "delete",
            ["path"] = pdfPath,
            ["outputPath"] = outputPath,
            ["pageIndex"] = 1
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.Contains("Deleted all 2 annotation(s)", result);
        using var resultDocument = new Document(outputPath);
        Assert.Empty(resultDocument.Pages[1].Annotations);
    }

    [Fact]
    public async Task DeleteAnnotation_NoAnnotations_ShouldThrow()
    {
        // Arrange
        var pdfPath = CreateTestPdf("test_delete_no_annotations.pdf");
        var arguments = new JsonObject
        {
            ["operation"] = "delete",
            ["path"] = pdfPath,
            ["pageIndex"] = 1
        };

        // Act & Assert
        await Assert.ThrowsAsync<ArgumentException>(() => _tool.ExecuteAsync(arguments));
    }

    [Fact]
    public async Task DeleteAnnotation_InvalidAnnotationIndex_ShouldThrow()
    {
        // Arrange
        var pdfPath = CreateTestPdf("test_delete_invalid_index.pdf");
        using (var document = new Document(pdfPath))
        {
            var page = document.Pages[1];
            page.Annotations.Add(new TextAnnotation(page, new Rectangle(100, 100, 200, 130)) { Contents = "Note" });
            document.Save(pdfPath);
        }

        var arguments = new JsonObject
        {
            ["operation"] = "delete",
            ["path"] = pdfPath,
            ["pageIndex"] = 1,
            ["annotationIndex"] = 99
        };

        // Act & Assert
        await Assert.ThrowsAsync<ArgumentException>(() => _tool.ExecuteAsync(arguments));
    }

    [Fact]
    public async Task EditAnnotation_ShouldModifyAnnotation()
    {
        // Arrange
        var pdfPath = CreateTestPdf("test_edit_annotation.pdf");
        using (var document = new Document(pdfPath))
        {
            var page = document.Pages[1];
            page.Annotations.Add(new TextAnnotation(page, new Rectangle(100, 100, 200, 130))
                { Contents = "Original Note" });
            document.Save(pdfPath);
        }

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
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.Contains("Edited annotation", result);
        using var resultDocument = new Document(outputPath);
        var editedAnnotation = resultDocument.Pages[1].Annotations[1] as TextAnnotation;
        Assert.NotNull(editedAnnotation);
        Assert.Equal("Updated Note", editedAnnotation.Contents);
    }

    [Fact]
    public async Task EditAnnotation_WithPosition_ShouldUpdatePosition()
    {
        // Arrange
        var pdfPath = CreateTestPdf("test_edit_position.pdf");
        using (var document = new Document(pdfPath))
        {
            var page = document.Pages[1];
            page.Annotations.Add(new TextAnnotation(page, new Rectangle(100, 100, 300, 150)) { Contents = "Note" });
            document.Save(pdfPath);
        }

        var outputPath = CreateTestFilePath("test_edit_position_output.pdf");
        var arguments = new JsonObject
        {
            ["operation"] = "edit",
            ["path"] = pdfPath,
            ["outputPath"] = outputPath,
            ["pageIndex"] = 1,
            ["annotationIndex"] = 1,
            ["text"] = "Moved Note",
            ["x"] = 200,
            ["y"] = 500
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.Contains("Edited annotation", result);
        using var resultDocument = new Document(outputPath);
        var editedAnnotation = resultDocument.Pages[1].Annotations[1] as TextAnnotation;
        Assert.NotNull(editedAnnotation);
        Assert.Equal(200, editedAnnotation.Rect.LLX, 1);
        Assert.Equal(500, editedAnnotation.Rect.LLY, 1);
    }

    [Fact]
    public async Task EditAnnotation_InvalidAnnotationIndex_ShouldThrow()
    {
        // Arrange
        var pdfPath = CreateTestPdf("test_edit_invalid_index.pdf");
        using (var document = new Document(pdfPath))
        {
            var page = document.Pages[1];
            page.Annotations.Add(new TextAnnotation(page, new Rectangle(100, 100, 200, 130)) { Contents = "Note" });
            document.Save(pdfPath);
        }

        var arguments = new JsonObject
        {
            ["operation"] = "edit",
            ["path"] = pdfPath,
            ["pageIndex"] = 1,
            ["annotationIndex"] = 99,
            ["text"] = "Test"
        };

        // Act & Assert
        await Assert.ThrowsAsync<ArgumentException>(() => _tool.ExecuteAsync(arguments));
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