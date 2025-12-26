using System.Drawing;
using System.Drawing.Imaging;
using System.Runtime.Versioning;
using Aspose.Words;
using Aspose.Words.Drawing;
using AsposeMcpServer.Tests.Helpers;
using AsposeMcpServer.Tools.Word;

namespace AsposeMcpServer.Tests.Word;

[SupportedOSPlatform("windows")]
public class WordImageToolTests : WordTestBase
{
    private readonly WordImageTool _tool = new();

    private string CreateTestImage(string fileName)
    {
        var imagePath = CreateTestFilePath(fileName);
        using var bitmap = new Bitmap(100, 100);
        using var graphics = Graphics.FromImage(bitmap);
        graphics.Clear(Color.Blue);
        bitmap.Save(imagePath, ImageFormat.Png);
        return imagePath;
    }

    [Fact]
    public async Task AddImage_ShouldAddImageToDocument()
    {
        // Arrange
        var docPath = CreateWordDocument("test_add_image.docx");
        var imagePath = CreateTestImage("test_image.png");
        var outputPath = CreateTestFilePath("test_add_image_output.docx");
        var arguments = CreateArguments("add", docPath, outputPath);
        arguments["imagePath"] = imagePath;
        arguments["width"] = 200;

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        var doc = new Document(outputPath);
        var shapes = doc.GetChildNodes(NodeType.Shape, true).Cast<Shape>().ToList();
        Assert.True(shapes.Count > 0, "Document should contain at least one image");
    }

    [Fact]
    public async Task EditImage_ShouldEditImageProperties()
    {
        // Arrange
        var docPath = CreateWordDocument("test_edit_image.docx");
        var imagePath = CreateTestImage("test_image2.png");
        var doc = new Document(docPath);
        var builder = new DocumentBuilder(doc);
        builder.InsertImage(imagePath);
        doc.Save(docPath);

        var outputPath = CreateTestFilePath("test_edit_image_output.docx");
        var arguments = CreateArguments("edit", docPath, outputPath);
        arguments["imageIndex"] = 0;
        arguments["width"] = 300;
        arguments["height"] = 200;

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        var resultDoc = new Document(outputPath);
        var shapes = resultDoc.GetChildNodes(NodeType.Shape, true).Cast<Shape>().ToList();
        Assert.True(shapes.Count > 0, "Document should contain image after editing");
    }

    [Fact]
    public async Task DeleteImage_ShouldDeleteImage()
    {
        // Arrange
        var docPath = CreateWordDocument("test_delete_image.docx");
        var imagePath = CreateTestImage("test_image3.png");
        var doc = new Document(docPath);
        var builder = new DocumentBuilder(doc);
        builder.InsertImage(imagePath);
        doc.Save(docPath);

        var shapesBefore = doc.GetChildNodes(NodeType.Shape, true).Cast<Shape>().ToList().Count;
        Assert.True(shapesBefore > 0, "Document should have image before deletion");

        var outputPath = CreateTestFilePath("test_delete_image_output.docx");
        var arguments = CreateArguments("delete", docPath, outputPath);
        arguments["imageIndex"] = 0;

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        var resultDoc = new Document(outputPath);
        var shapesAfter = resultDoc.GetChildNodes(NodeType.Shape, true).Cast<Shape>().ToList().Count;
        Assert.True(shapesAfter < shapesBefore,
            $"Image should be deleted. Before: {shapesBefore}, After: {shapesAfter}");
    }

    [Fact]
    public async Task GetImages_ShouldReturnAllImages()
    {
        // Arrange
        var docPath = CreateWordDocument("test_get_images.docx");
        var imagePath = CreateTestImage("test_image4.png");
        var doc = new Document(docPath);
        var builder = new DocumentBuilder(doc);
        builder.InsertImage(imagePath);
        doc.Save(docPath);

        var arguments = CreateArguments("get", docPath);
        arguments["operation"] = "get";

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.NotNull(result);
        Assert.NotEmpty(result);
        Assert.Contains("Image", result, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public async Task ReplaceImage_ShouldReplaceImage()
    {
        // Arrange
        var docPath = CreateWordDocument("test_replace_image.docx");
        var originalImagePath = CreateTestImage("test_image5.png");
        var doc = new Document(docPath);
        var builder = new DocumentBuilder(doc);
        builder.InsertImage(originalImagePath);
        doc.Save(docPath);

        var newImagePath = CreateTestImage("test_image6.png");
        var outputPath = CreateTestFilePath("test_replace_image_output.docx");
        var arguments = CreateArguments("replace", docPath, outputPath);
        arguments["imageIndex"] = 0;
        arguments["newImagePath"] = newImagePath;

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        var resultDoc = new Document(outputPath);
        var shapes = resultDoc.GetChildNodes(NodeType.Shape, true).Cast<Shape>().ToList();
        Assert.True(shapes.Count > 0, "Document should contain image after replacement");
    }

    [Fact]
    public async Task ExtractImages_ShouldExtractAllImages()
    {
        // Arrange
        var docPath = CreateWordDocument("test_extract_images.docx");
        var imagePath = CreateTestImage("test_image7.png");
        var doc = new Document(docPath);
        var builder = new DocumentBuilder(doc);
        builder.InsertImage(imagePath);
        doc.Save(docPath);

        var outputDir = CreateTestFilePath("extract_images_output");
        Directory.CreateDirectory(outputDir);
        var arguments = CreateArguments("extract", docPath);
        arguments["outputDir"] = outputDir;

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        var files = Directory.GetFiles(outputDir);
        Assert.True(files.Length > 0, "Should extract at least one image");
    }

    [Fact]
    public async Task AddImage_WithHyperlink_ShouldSetImageHyperlink()
    {
        // Arrange
        var docPath = CreateWordDocument("test_add_image_hyperlink.docx");
        var imagePath = CreateTestImage("test_image_hyperlink.png");
        var outputPath = CreateTestFilePath("test_add_image_hyperlink_output.docx");
        var testUrl = "https://example.com/test";
        var arguments = CreateArguments("add", docPath, outputPath);
        arguments["imagePath"] = imagePath;
        arguments["linkUrl"] = testUrl;
        arguments["alternativeText"] = "Test alt text";
        arguments["title"] = "Test title";

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.Contains("Hyperlink:", result);
        Assert.Contains(testUrl, result);

        var doc = new Document(outputPath);
        var shapes = doc.GetChildNodes(NodeType.Shape, true).Cast<Shape>().Where(s => s.HasImage).ToList();
        Assert.True(shapes.Count > 0, "Document should contain at least one image");
        Assert.Equal(testUrl, shapes[0].HRef);
        Assert.Equal("Test alt text", shapes[0].AlternativeText);
        Assert.Equal("Test title", shapes[0].Title);
    }

    [Fact]
    public async Task EditImage_WithHyperlink_ShouldUpdateImageHyperlink()
    {
        // Arrange
        var docPath = CreateWordDocument("test_edit_image_hyperlink.docx");
        var imagePath = CreateTestImage("test_image_hyperlink2.png");
        var doc = new Document(docPath);
        var builder = new DocumentBuilder(doc);
        builder.InsertImage(imagePath);
        doc.Save(docPath);

        var outputPath = CreateTestFilePath("test_edit_image_hyperlink_output.docx");
        var testUrl = "https://example.com/updated";
        var arguments = CreateArguments("edit", docPath, outputPath);
        arguments["imageIndex"] = 0;
        arguments["linkUrl"] = testUrl;

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.Contains("Hyperlink:", result);

        var resultDoc = new Document(outputPath);
        var shapes = resultDoc.GetChildNodes(NodeType.Shape, true).Cast<Shape>().Where(s => s.HasImage).ToList();
        Assert.True(shapes.Count > 0, "Document should contain image after editing");
        Assert.Equal(testUrl, shapes[0].HRef);
    }

    [Fact]
    public async Task EditImage_RemoveHyperlink_ShouldClearImageHyperlink()
    {
        // Arrange
        var docPath = CreateWordDocument("test_remove_image_hyperlink.docx");
        var imagePath = CreateTestImage("test_image_hyperlink3.png");
        var doc = new Document(docPath);
        var builder = new DocumentBuilder(doc);
        var shape = builder.InsertImage(imagePath);
        shape.HRef = "https://original.com";
        doc.Save(docPath);

        var outputPath = CreateTestFilePath("test_remove_image_hyperlink_output.docx");
        var arguments = CreateArguments("edit", docPath, outputPath);
        arguments["imageIndex"] = 0;
        arguments["linkUrl"] = ""; // Empty string to remove hyperlink

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.Contains("Hyperlink: removed", result);

        var resultDoc = new Document(outputPath);
        var shapes = resultDoc.GetChildNodes(NodeType.Shape, true).Cast<Shape>().Where(s => s.HasImage).ToList();
        Assert.True(shapes.Count > 0, "Document should contain image after editing");
        Assert.True(string.IsNullOrEmpty(shapes[0].HRef), "Hyperlink should be removed");
    }

    [Fact]
    public async Task GetImages_WithHyperlink_ShouldShowHyperlinkInfo()
    {
        // Arrange
        var docPath = CreateWordDocument("test_get_images_hyperlink.docx");
        var imagePath = CreateTestImage("test_image_hyperlink4.png");
        var testUrl = "https://example.com/gettest";
        var doc = new Document(docPath);
        var builder = new DocumentBuilder(doc);
        var shape = builder.InsertImage(imagePath);
        shape.HRef = testUrl;
        shape.AlternativeText = "Alt text for test";
        doc.Save(docPath);

        var arguments = CreateArguments("get", docPath);

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.Contains("\"hyperlink\"", result); // JSON format
        Assert.Contains(testUrl, result);
        Assert.Contains("\"altText\"", result); // JSON format
    }
}