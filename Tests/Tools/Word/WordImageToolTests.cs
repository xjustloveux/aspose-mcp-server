using System.Drawing;
using System.Drawing.Imaging;
using System.Runtime.Versioning;
using Aspose.Words;
using Aspose.Words.Drawing;
using AsposeMcpServer.Tests.Helpers;
using AsposeMcpServer.Tools.Word;

namespace AsposeMcpServer.Tests.Tools.Word;

/// <summary>
///     Integration tests for WordImageTool.
///     Focuses on session management, file I/O, and operation routing.
///     Detailed parameter validation and business logic tests are in Handler tests.
/// </summary>
[SupportedOSPlatform("windows")]
public class WordImageToolTests : WordTestBase
{
    private readonly WordImageTool _tool;

    public WordImageToolTests()
    {
        _tool = new WordImageTool(SessionManager);
    }

    private string CreateTestImage(string fileName)
    {
        var imagePath = CreateTestFilePath(fileName);
        using var bitmap = new Bitmap(100, 100);
        using var graphics = Graphics.FromImage(bitmap);
        graphics.Clear(Color.Blue);
        bitmap.Save(imagePath, ImageFormat.Png);
        return imagePath;
    }

    #region File I/O Smoke Tests

    [Fact]
    public void AddImage_ShouldAddImageAndPersistToFile()
    {
        var docPath = CreateWordDocument("test_add_image.docx");
        var imagePath = CreateTestImage("test_image.png");
        var outputPath = CreateTestFilePath("test_add_image_output.docx");
        _tool.Execute("add", docPath, outputPath: outputPath, imagePath: imagePath, width: 200);
        var doc = new Document(outputPath);
        var shapes = doc.GetChildNodes(NodeType.Shape, true).Cast<Shape>().ToList();
        Assert.True(shapes.Count > 0);
    }

    [Fact]
    public void GetImages_ShouldReturnImagesFromFile()
    {
        var docPath = CreateWordDocument("test_get_images.docx");
        var imagePath = CreateTestImage("test_image_get.png");
        var doc = new Document(docPath);
        var builder = new DocumentBuilder(doc);
        builder.InsertImage(imagePath);
        doc.Save(docPath);
        var result = _tool.Execute("get", docPath);
        Assert.Contains("Image", result, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void EditImage_ShouldEditImageAndPersistToFile()
    {
        var docPath = CreateWordDocument("test_edit_image.docx");
        var imagePath = CreateTestImage("test_image_edit.png");
        var doc = new Document(docPath);
        var builder = new DocumentBuilder(doc);
        builder.InsertImage(imagePath);
        doc.Save(docPath);

        var outputPath = CreateTestFilePath("test_edit_image_output.docx");
        _tool.Execute("edit", docPath, outputPath: outputPath, imageIndex: 0, width: 300, height: 200);
        Assert.True(File.Exists(outputPath));
    }

    [Fact]
    public void DeleteImage_ShouldDeleteImageAndPersistToFile()
    {
        var docPath = CreateWordDocument("test_delete_image.docx");
        var imagePath = CreateTestImage("test_image_delete.png");
        var doc = new Document(docPath);
        var builder = new DocumentBuilder(doc);
        builder.InsertImage(imagePath);
        doc.Save(docPath);

        var shapesBefore = doc.GetChildNodes(NodeType.Shape, true).Cast<Shape>().ToList().Count;
        var outputPath = CreateTestFilePath("test_delete_image_output.docx");
        _tool.Execute("delete", docPath, outputPath: outputPath, imageIndex: 0);
        var resultDoc = new Document(outputPath);
        var shapesAfter = resultDoc.GetChildNodes(NodeType.Shape, true).Cast<Shape>().ToList().Count;
        Assert.True(shapesAfter < shapesBefore);
    }

    [Fact]
    public void ReplaceImage_ShouldReplaceImageAndPersistToFile()
    {
        var docPath = CreateWordDocument("test_replace_image.docx");
        var originalImagePath = CreateTestImage("test_image_original.png");
        var doc = new Document(docPath);
        var builder = new DocumentBuilder(doc);
        builder.InsertImage(originalImagePath);
        doc.Save(docPath);

        var newImagePath = CreateTestImage("test_image_new.png");
        var outputPath = CreateTestFilePath("test_replace_image_output.docx");
        _tool.Execute("replace", docPath, outputPath: outputPath, imageIndex: 0, newImagePath: newImagePath);
        Assert.True(File.Exists(outputPath));
    }

    [Fact]
    public void ExtractImages_ShouldExtractImagesFromFile()
    {
        var docPath = CreateWordDocument("test_extract_images.docx");
        var imagePath = CreateTestImage("test_image_extract.png");
        var doc = new Document(docPath);
        var builder = new DocumentBuilder(doc);
        builder.InsertImage(imagePath);
        doc.Save(docPath);

        var outputDir = CreateTestFilePath("extract_images_output");
        Directory.CreateDirectory(outputDir);
        _tool.Execute("extract", docPath, outputDir: outputDir);
        var files = Directory.GetFiles(outputDir);
        Assert.True(files.Length > 0);
    }

    #endregion

    #region Operation Routing

    [Theory]
    [InlineData("ADD")]
    [InlineData("Add")]
    [InlineData("add")]
    public void Operation_ShouldBeCaseInsensitive(string operation)
    {
        var docPath = CreateWordDocument($"test_case_{operation}.docx");
        var imagePath = CreateTestImage($"test_case_{operation}.png");
        var outputPath = CreateTestFilePath($"test_case_{operation}_output.docx");
        var result = _tool.Execute(operation, docPath, outputPath: outputPath, imagePath: imagePath);
        Assert.StartsWith("Image added", result);
    }

    [Fact]
    public void Execute_WithUnknownOperation_ShouldThrowArgumentException()
    {
        var docPath = CreateWordDocumentWithContent("test_unknown_op.docx", "Test content");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("unknown_operation", docPath));
        Assert.Contains("Unknown operation", ex.Message);
    }

    [Fact]
    public void Execute_WithNoPathOrSessionId_ShouldThrowException()
    {
        Assert.ThrowsAny<Exception>(() => _tool.Execute("get"));
    }

    #endregion

    #region Session Management

    [Fact]
    public void GetImages_WithSessionId_ShouldReturnImages()
    {
        var docPath = CreateWordDocument("test_session_get_img.docx");
        var imagePath = CreateTestImage("test_image_session_get.png");
        var doc = new Document(docPath);
        var builder = new DocumentBuilder(doc);
        builder.InsertImage(imagePath);
        doc.Save(docPath);

        var sessionId = OpenSession(docPath);
        var result = _tool.Execute("get", sessionId: sessionId);
        Assert.Contains("Image", result, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void AddImage_WithSessionId_ShouldAddImageInMemory()
    {
        var docPath = CreateWordDocument("test_session_add_img.docx");
        var imagePath = CreateTestImage("test_image_session_add.png");
        var sessionId = OpenSession(docPath);
        var result = _tool.Execute("add", sessionId: sessionId, imagePath: imagePath, width: 150);
        Assert.Contains("added", result, StringComparison.OrdinalIgnoreCase);

        var doc = SessionManager.GetDocument<Document>(sessionId);
        var shapes = doc.GetChildNodes(NodeType.Shape, true).Cast<Shape>().Where(s => s.HasImage).ToList();
        Assert.True(shapes.Count > 0);
    }

    [Fact]
    public void EditImage_WithSessionId_ShouldEditImageInMemory()
    {
        var docPath = CreateWordDocument("test_session_edit_img.docx");
        var imagePath = CreateTestImage("test_image_session_edit.png");
        var doc = new Document(docPath);
        var builder = new DocumentBuilder(doc);
        builder.InsertImage(imagePath);
        doc.Save(docPath);

        var sessionId = OpenSession(docPath);
        var result = _tool.Execute("edit", sessionId: sessionId, imageIndex: 0, width: 250, height: 180);
        Assert.Contains("Image", result);
    }

    [Fact]
    public void DeleteImage_WithSessionId_ShouldDeleteImageInMemory()
    {
        var docPath = CreateWordDocument("test_session_delete_img.docx");
        var imagePath = CreateTestImage("test_image_session_delete.png");
        var doc = new Document(docPath);
        var builder = new DocumentBuilder(doc);
        builder.InsertImage(imagePath);
        doc.Save(docPath);

        var sessionId = OpenSession(docPath);
        var docBefore = SessionManager.GetDocument<Document>(sessionId);
        var shapesBefore = docBefore.GetChildNodes(NodeType.Shape, true).Cast<Shape>().Where(s => s.HasImage).ToList();
        Assert.True(shapesBefore.Count > 0);

        _tool.Execute("delete", sessionId: sessionId, imageIndex: 0);
        var sessionDoc = SessionManager.GetDocument<Document>(sessionId);
        var shapesAfter = sessionDoc.GetChildNodes(NodeType.Shape, true).Cast<Shape>().Where(s => s.HasImage).ToList();
        Assert.True(shapesAfter.Count < shapesBefore.Count);
    }

    [Fact]
    public void ReplaceImage_WithSessionId_ShouldReplaceInMemory()
    {
        var docPath = CreateWordDocument("test_session_replace_img.docx");
        var imagePath = CreateTestImage("test_session_replace.png");
        var doc = new Document(docPath);
        var builder = new DocumentBuilder(doc);
        builder.InsertImage(imagePath);
        doc.Save(docPath);

        var sessionId = OpenSession(docPath);
        var newImagePath = CreateTestImage("test_session_replace_new.png");
        var result = _tool.Execute("replace", sessionId: sessionId, imageIndex: 0, newImagePath: newImagePath);
        Assert.StartsWith("Image #0 replaced", result);
    }

    [Fact]
    public void Execute_WithInvalidSessionId_ShouldThrowKeyNotFoundException()
    {
        Assert.Throws<KeyNotFoundException>(() =>
            _tool.Execute("get", sessionId: "invalid_session_id"));
    }

    [Fact]
    public void Execute_WithBothPathAndSessionId_ShouldPreferSessionId()
    {
        var docPath1 = CreateWordDocument("test_path_img.docx");
        var imagePath1 = CreateTestImage("test_path_image.png");
        var doc1 = new Document(docPath1);
        var builder1 = new DocumentBuilder(doc1);
        var shape1 = builder1.InsertImage(imagePath1);
        shape1.AlternativeText = "Path Image Alt";
        doc1.Save(docPath1);

        var docPath2 = CreateWordDocument("test_session_img.docx");
        var imagePath2 = CreateTestImage("test_session_image.png");
        var doc2 = new Document(docPath2);
        var builder2 = new DocumentBuilder(doc2);
        var shape2 = builder2.InsertImage(imagePath2);
        shape2.AlternativeText = "Session Image Alt Unique";
        doc2.Save(docPath2);

        var sessionId = OpenSession(docPath2);
        var result = _tool.Execute("get", docPath1, sessionId);

        Assert.Contains("Session Image Alt Unique", result);
        Assert.DoesNotContain("Path Image Alt", result);
    }

    #endregion
}
