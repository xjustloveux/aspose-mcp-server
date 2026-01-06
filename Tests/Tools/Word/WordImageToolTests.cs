using System.Drawing;
using System.Drawing.Imaging;
using System.Runtime.Versioning;
using Aspose.Words;
using Aspose.Words.Drawing;
using AsposeMcpServer.Tests.Helpers;
using AsposeMcpServer.Tools.Word;

namespace AsposeMcpServer.Tests.Tools.Word;

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

    #region General

    [Fact]
    public void AddImage_ShouldAddImageToDocument()
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
    public void EditImage_ShouldEditImageProperties()
    {
        var docPath = CreateWordDocument("test_edit_image.docx");
        var imagePath = CreateTestImage("test_image2.png");
        var doc = new Document(docPath);
        var builder = new DocumentBuilder(doc);
        builder.InsertImage(imagePath);
        doc.Save(docPath);

        var outputPath = CreateTestFilePath("test_edit_image_output.docx");
        _tool.Execute("edit", docPath, outputPath: outputPath, imageIndex: 0, width: 300, height: 200);
        var resultDoc = new Document(outputPath);
        var shapes = resultDoc.GetChildNodes(NodeType.Shape, true).Cast<Shape>().ToList();
        Assert.True(shapes.Count > 0);
    }

    [Fact]
    public void DeleteImage_ShouldDeleteImage()
    {
        var docPath = CreateWordDocument("test_delete_image.docx");
        var imagePath = CreateTestImage("test_image3.png");
        var doc = new Document(docPath);
        var builder = new DocumentBuilder(doc);
        builder.InsertImage(imagePath);
        doc.Save(docPath);

        var shapesBefore = doc.GetChildNodes(NodeType.Shape, true).Cast<Shape>().ToList().Count;
        Assert.True(shapesBefore > 0);

        var outputPath = CreateTestFilePath("test_delete_image_output.docx");
        _tool.Execute("delete", docPath, outputPath: outputPath, imageIndex: 0);
        var resultDoc = new Document(outputPath);
        var shapesAfter = resultDoc.GetChildNodes(NodeType.Shape, true).Cast<Shape>().ToList().Count;
        Assert.True(shapesAfter < shapesBefore);
    }

    [Fact]
    public void GetImages_ShouldReturnAllImages()
    {
        var docPath = CreateWordDocument("test_get_images.docx");
        var imagePath = CreateTestImage("test_image4.png");
        var doc = new Document(docPath);
        var builder = new DocumentBuilder(doc);
        builder.InsertImage(imagePath);
        doc.Save(docPath);
        var result = _tool.Execute("get", docPath);
        Assert.NotNull(result);
        Assert.NotEmpty(result);
        Assert.Contains("Image", result, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void ReplaceImage_ShouldReplaceImage()
    {
        var docPath = CreateWordDocument("test_replace_image.docx");
        var originalImagePath = CreateTestImage("test_image5.png");
        var doc = new Document(docPath);
        var builder = new DocumentBuilder(doc);
        builder.InsertImage(originalImagePath);
        doc.Save(docPath);

        var newImagePath = CreateTestImage("test_image6.png");
        var outputPath = CreateTestFilePath("test_replace_image_output.docx");
        _tool.Execute("replace", docPath, outputPath: outputPath, imageIndex: 0, newImagePath: newImagePath);
        var resultDoc = new Document(outputPath);
        var shapes = resultDoc.GetChildNodes(NodeType.Shape, true).Cast<Shape>().ToList();
        Assert.True(shapes.Count > 0);
    }

    [Fact]
    public void ExtractImages_ShouldExtractAllImages()
    {
        var docPath = CreateWordDocument("test_extract_images.docx");
        var imagePath = CreateTestImage("test_image7.png");
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

    [Fact]
    public void AddImage_WithHyperlink_ShouldSetImageHyperlink()
    {
        var docPath = CreateWordDocument("test_add_image_hyperlink.docx");
        var imagePath = CreateTestImage("test_image_hyperlink.png");
        var outputPath = CreateTestFilePath("test_add_image_hyperlink_output.docx");
        var testUrl = "https://example.com/test";
        var result = _tool.Execute("add", docPath, outputPath: outputPath,
            imagePath: imagePath, linkUrl: testUrl, alternativeText: "Test alt text", title: "Test title");
        Assert.Contains("Hyperlink:", result);
        Assert.Contains(testUrl, result);

        var doc = new Document(outputPath);
        var shapes = doc.GetChildNodes(NodeType.Shape, true).Cast<Shape>().Where(s => s.HasImage).ToList();
        Assert.True(shapes.Count > 0);
        Assert.Equal(testUrl, shapes[0].HRef);
        Assert.Equal("Test alt text", shapes[0].AlternativeText);
        Assert.Equal("Test title", shapes[0].Title);
    }

    [Fact]
    public void EditImage_WithHyperlink_ShouldUpdateImageHyperlink()
    {
        var docPath = CreateWordDocument("test_edit_image_hyperlink.docx");
        var imagePath = CreateTestImage("test_image_hyperlink2.png");
        var doc = new Document(docPath);
        var builder = new DocumentBuilder(doc);
        builder.InsertImage(imagePath);
        doc.Save(docPath);

        var outputPath = CreateTestFilePath("test_edit_image_hyperlink_output.docx");
        var testUrl = "https://example.com/updated";
        var result = _tool.Execute("edit", docPath, outputPath: outputPath, imageIndex: 0, linkUrl: testUrl);
        Assert.Contains("Hyperlink:", result);

        var resultDoc = new Document(outputPath);
        var shapes = resultDoc.GetChildNodes(NodeType.Shape, true).Cast<Shape>().Where(s => s.HasImage).ToList();
        Assert.True(shapes.Count > 0);
        Assert.Equal(testUrl, shapes[0].HRef);
    }

    [Fact]
    public void EditImage_RemoveHyperlink_ShouldClearImageHyperlink()
    {
        var docPath = CreateWordDocument("test_remove_image_hyperlink.docx");
        var imagePath = CreateTestImage("test_image_hyperlink3.png");
        var doc = new Document(docPath);
        var builder = new DocumentBuilder(doc);
        var shape = builder.InsertImage(imagePath);
        shape.HRef = "https://original.com";
        doc.Save(docPath);

        var outputPath = CreateTestFilePath("test_remove_image_hyperlink_output.docx");
        var result = _tool.Execute("edit", docPath, outputPath: outputPath, imageIndex: 0, linkUrl: "");
        Assert.Contains("Hyperlink: removed", result);

        var resultDoc = new Document(outputPath);
        var shapes = resultDoc.GetChildNodes(NodeType.Shape, true).Cast<Shape>().Where(s => s.HasImage).ToList();
        Assert.True(shapes.Count > 0);
        Assert.True(string.IsNullOrEmpty(shapes[0].HRef));
    }

    [Fact]
    public void GetImages_WithHyperlink_ShouldShowHyperlinkInfo()
    {
        var docPath = CreateWordDocument("test_get_images_hyperlink.docx");
        var imagePath = CreateTestImage("test_image_hyperlink4.png");
        var testUrl = "https://example.com/gettest";
        var doc = new Document(docPath);
        var builder = new DocumentBuilder(doc);
        var shape = builder.InsertImage(imagePath);
        shape.HRef = testUrl;
        shape.AlternativeText = "Alt text for test";
        doc.Save(docPath);
        var result = _tool.Execute("get", docPath);
        Assert.Contains("\"hyperlink\"", result);
        Assert.Contains(testUrl, result);
        Assert.Contains("\"altText\"", result);
    }

    [Theory]
    [InlineData("ADD")]
    [InlineData("Add")]
    [InlineData("add")]
    public void Operation_ShouldBeCaseInsensitive_Add(string operation)
    {
        var docPath = CreateWordDocument($"test_case_{operation}.docx");
        var imagePath = CreateTestImage($"test_case_{operation}.png");
        var outputPath = CreateTestFilePath($"test_case_{operation}_output.docx");
        var result = _tool.Execute(operation, docPath, outputPath: outputPath, imagePath: imagePath);
        Assert.StartsWith("Image added", result);
    }

    [Theory]
    [InlineData("GET")]
    [InlineData("Get")]
    [InlineData("get")]
    public void Operation_ShouldBeCaseInsensitive_Get(string operation)
    {
        var docPath = CreateWordDocument($"test_case_get_{operation}.docx");
        var imagePath = CreateTestImage($"test_case_get_{operation}.png");
        var doc = new Document(docPath);
        var builder = new DocumentBuilder(doc);
        builder.InsertImage(imagePath);
        doc.Save(docPath);
        var result = _tool.Execute(operation, docPath);
        Assert.Contains("images", result.ToLower());
    }

    [Theory]
    [InlineData("EDIT")]
    [InlineData("Edit")]
    [InlineData("edit")]
    public void Operation_ShouldBeCaseInsensitive_Edit(string operation)
    {
        var docPath = CreateWordDocument($"test_case_edit_{operation}.docx");
        var imagePath = CreateTestImage($"test_case_edit_{operation}.png");
        var doc = new Document(docPath);
        var builder = new DocumentBuilder(doc);
        builder.InsertImage(imagePath);
        doc.Save(docPath);
        var outputPath = CreateTestFilePath($"test_case_edit_{operation}_output.docx");
        var result = _tool.Execute(operation, docPath, outputPath: outputPath, imageIndex: 0, width: 100);
        Assert.StartsWith("Image 0 edited", result);
    }

    [Theory]
    [InlineData("DELETE")]
    [InlineData("Delete")]
    [InlineData("delete")]
    public void Operation_ShouldBeCaseInsensitive_Delete(string operation)
    {
        var docPath = CreateWordDocument($"test_case_delete_{operation}.docx");
        var imagePath = CreateTestImage($"test_case_delete_{operation}.png");
        var doc = new Document(docPath);
        var builder = new DocumentBuilder(doc);
        builder.InsertImage(imagePath);
        doc.Save(docPath);
        var outputPath = CreateTestFilePath($"test_case_delete_{operation}_output.docx");
        var result = _tool.Execute(operation, docPath, outputPath: outputPath, imageIndex: 0);
        Assert.StartsWith("Image #0", result);
    }

    [Theory]
    [InlineData("REPLACE")]
    [InlineData("Replace")]
    [InlineData("replace")]
    public void Operation_ShouldBeCaseInsensitive_Replace(string operation)
    {
        var docPath = CreateWordDocument($"test_case_replace_{operation}.docx");
        var imagePath = CreateTestImage($"test_case_replace_{operation}.png");
        var doc = new Document(docPath);
        var builder = new DocumentBuilder(doc);
        builder.InsertImage(imagePath);
        doc.Save(docPath);
        var newImagePath = CreateTestImage($"test_case_replace_new_{operation}.png");
        var outputPath = CreateTestFilePath($"test_case_replace_{operation}_output.docx");
        var result = _tool.Execute(operation, docPath, outputPath: outputPath, imageIndex: 0,
            newImagePath: newImagePath);
        Assert.StartsWith("Image #0 replaced", result);
    }

    [Theory]
    [InlineData("EXTRACT")]
    [InlineData("Extract")]
    [InlineData("extract")]
    public void Operation_ShouldBeCaseInsensitive_Extract(string operation)
    {
        var docPath = CreateWordDocument($"test_case_extract_{operation}.docx");
        var imagePath = CreateTestImage($"test_case_extract_{operation}.png");
        var doc = new Document(docPath);
        var builder = new DocumentBuilder(doc);
        builder.InsertImage(imagePath);
        doc.Save(docPath);
        var outputDir = CreateTestFilePath($"test_case_extract_{operation}_output");
        Directory.CreateDirectory(outputDir);
        var result = _tool.Execute(operation, docPath, outputDir: outputDir);
        Assert.Contains("extracted", result.ToLower());
    }

    [Theory]
    [InlineData("LEFT")]
    [InlineData("Left")]
    [InlineData("left")]
    public void Alignment_ShouldBeCaseInsensitive(string alignment)
    {
        var docPath = CreateWordDocument($"test_align_{alignment}.docx");
        var imagePath = CreateTestImage($"test_align_{alignment}.png");
        var outputPath = CreateTestFilePath($"test_align_{alignment}_output.docx");
        var result = _tool.Execute("add", docPath, outputPath: outputPath, imagePath: imagePath, alignment: alignment);
        Assert.StartsWith("Image added", result);
        Assert.Contains("alignment:", result.ToLower());
    }

    [Theory]
    [InlineData("INLINE")]
    [InlineData("Inline")]
    [InlineData("inline")]
    [InlineData("SQUARE")]
    [InlineData("Square")]
    [InlineData("square")]
    public void TextWrapping_ShouldBeCaseInsensitive(string textWrapping)
    {
        var docPath = CreateWordDocument($"test_wrap_{textWrapping}.docx");
        var imagePath = CreateTestImage($"test_wrap_{textWrapping}.png");
        var outputPath = CreateTestFilePath($"test_wrap_{textWrapping}_output.docx");
        var result = _tool.Execute("add", docPath, outputPath: outputPath, imagePath: imagePath,
            textWrapping: textWrapping);
        Assert.StartsWith("Image added", result);
        Assert.Contains("text wrapping:", result.ToLower());
    }

    [Theory]
    [InlineData("CENTER")]
    [InlineData("Center")]
    [InlineData("center")]
    [InlineData("RIGHT")]
    [InlineData("Right")]
    [InlineData("right")]
    public void HorizontalAlignment_ShouldBeCaseInsensitive(string alignment)
    {
        var docPath = CreateWordDocument($"test_horiz_{alignment}.docx");
        var imagePath = CreateTestImage($"test_horiz_{alignment}.png");
        var doc = new Document(docPath);
        var builder = new DocumentBuilder(doc);
        builder.InsertImage(imagePath);
        doc.Save(docPath);
        var outputPath = CreateTestFilePath($"test_horiz_{alignment}_output.docx");
        var result = _tool.Execute("edit", docPath, outputPath: outputPath, imageIndex: 0,
            textWrapping: "square", horizontalAlignment: alignment);
        Assert.StartsWith("Image 0 edited", result);
    }

    [Fact]
    public void AddImage_WithCaption_ShouldAddCaptionBelow()
    {
        var docPath = CreateWordDocument("test_add_caption.docx");
        var imagePath = CreateTestImage("test_caption.png");
        var outputPath = CreateTestFilePath("test_add_caption_output.docx");
        var result = _tool.Execute("add", docPath, outputPath: outputPath,
            imagePath: imagePath, caption: "Test Caption", captionPosition: "below");
        Assert.Contains("Caption:", result);
        Assert.Contains("Test Caption", result);
        Assert.True(File.Exists(outputPath));
    }

    [Fact]
    public void AddImage_WithCaption_ShouldAddCaptionAbove()
    {
        var docPath = CreateWordDocument("test_add_caption_above.docx");
        var imagePath = CreateTestImage("test_caption_above.png");
        var outputPath = CreateTestFilePath("test_add_caption_above_output.docx");
        var result = _tool.Execute("add", docPath, outputPath: outputPath,
            imagePath: imagePath, caption: "Caption Above", captionPosition: "above");
        Assert.Contains("Caption:", result);
        Assert.Contains("above", result);
        Assert.True(File.Exists(outputPath));
    }

    [SkippableFact]
    public void GetImages_WithNoImages_ShouldReturnEmptyResult()
    {
        SkipInEvaluationMode(AsposeLibraryType.Words, "Evaluation watermark is treated as an image");
        var docPath = CreateWordDocumentWithContent("test_get_no_images.docx", "Text only");
        var result = _tool.Execute("get", docPath);
        Assert.Contains("\"count\": 0", result);
        Assert.Contains("No images found", result);
    }

    [Fact]
    public void ExtractImages_WithSpecificIndex_ShouldExtractSingleImage()
    {
        var docPath = CreateWordDocument("test_extract_single.docx");
        var imagePath1 = CreateTestImage("test_extract_single1.png");
        var imagePath2 = CreateTestImage("test_extract_single2.png");
        var doc = new Document(docPath);
        var builder = new DocumentBuilder(doc);
        builder.InsertImage(imagePath1);
        builder.InsertImage(imagePath2);
        doc.Save(docPath);

        var outputDir = CreateTestFilePath("extract_single_output");
        Directory.CreateDirectory(outputDir);
        var result = _tool.Execute("extract", docPath, outputDir: outputDir, extractImageIndex: 0);
        Assert.Contains("image #0", result.ToLower());
        var files = Directory.GetFiles(outputDir);
        Assert.Single(files);
    }

    [Fact]
    public void ReplaceImage_WithSmartFit_ShouldPreserveWidthAndCalculateHeight()
    {
        var docPath = CreateWordDocument("test_replace_smartfit.docx");
        var originalImagePath = CreateTestImage("test_smartfit_original.png");
        var doc = new Document(docPath);
        var builder = new DocumentBuilder(doc);
        builder.InsertImage(originalImagePath);
        doc.Save(docPath);

        var newImagePath = CreateTestImage("test_smartfit_new.png");
        var outputPath = CreateTestFilePath("test_replace_smartfit_output.docx");
        var result = _tool.Execute("replace", docPath, outputPath: outputPath,
            imageIndex: 0, newImagePath: newImagePath, preserveSize: true, smartFit: true);
        Assert.StartsWith("Image #0 replaced", result);
        Assert.Contains("smart fit", result.ToLower());
    }

    #endregion

    #region Exception

    [Fact]
    public void Execute_WithUnknownOperation_ShouldThrowArgumentException()
    {
        var docPath = CreateWordDocumentWithContent("test_unknown_op.docx", "Test content");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("unknown_operation", docPath));

        Assert.Contains("Unknown operation", ex.Message);
        Assert.Contains("unknown_operation", ex.Message);
    }

    [Fact]
    public void AddImage_WithMissingImagePath_ShouldThrowFileNotFoundException()
    {
        var docPath = CreateWordDocument("test_add_missing_image.docx");
        var outputPath = CreateTestFilePath("test_add_missing_image_output.docx");

        Assert.Throws<FileNotFoundException>(() =>
            _tool.Execute("add", docPath, outputPath: outputPath, imagePath: null));
    }

    [Fact]
    public void AddImage_WithNonExistentImagePath_ShouldThrowFileNotFoundException()
    {
        var docPath = CreateWordDocument("test_add_nonexistent_image.docx");
        var outputPath = CreateTestFilePath("test_add_nonexistent_image_output.docx");
        Assert.Throws<FileNotFoundException>(() =>
            _tool.Execute("add", docPath, outputPath: outputPath,
                imagePath: "C:\\nonexistent\\image.png"));
    }

    [Fact]
    public void EditImage_WithInvalidIndex_ShouldThrowArgumentException()
    {
        var docPath = CreateWordDocument("test_edit_invalid_index.docx");
        var imagePath = CreateTestImage("test_image_invalid_idx.png");
        var doc = new Document(docPath);
        var builder = new DocumentBuilder(doc);
        builder.InsertImage(imagePath);
        doc.Save(docPath);

        var outputPath = CreateTestFilePath("test_edit_invalid_index_output.docx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("edit", docPath, outputPath: outputPath, imageIndex: 999, width: 100));

        Assert.Contains("out of range", ex.Message);
    }

    [Fact]
    public void DeleteImage_WithInvalidIndex_ShouldThrowArgumentException()
    {
        var docPath = CreateWordDocument("test_delete_invalid_index.docx");
        var outputPath = CreateTestFilePath("test_delete_invalid_index_output.docx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("delete", docPath, outputPath: outputPath, imageIndex: 0));

        Assert.Contains("out of range", ex.Message);
    }

    [Fact]
    public void ReplaceImage_WithInvalidIndex_ShouldThrowArgumentException()
    {
        var docPath = CreateWordDocument("test_replace_invalid_index.docx");
        var imagePath = CreateTestImage("test_replace_invalid.png");
        var doc = new Document(docPath);
        var builder = new DocumentBuilder(doc);
        builder.InsertImage(imagePath);
        doc.Save(docPath);

        var newImagePath = CreateTestImage("test_replace_new_invalid.png");
        var outputPath = CreateTestFilePath("test_replace_invalid_index_output.docx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("replace", docPath, outputPath: outputPath, imageIndex: 999, newImagePath: newImagePath));

        Assert.Contains("out of range", ex.Message);
    }

    [Fact]
    public void ReplaceImage_WithNonExistentNewImage_ShouldThrowFileNotFoundException()
    {
        var docPath = CreateWordDocument("test_replace_nonexistent.docx");
        var imagePath = CreateTestImage("test_replace_nonexistent.png");
        var doc = new Document(docPath);
        var builder = new DocumentBuilder(doc);
        builder.InsertImage(imagePath);
        doc.Save(docPath);

        var outputPath = CreateTestFilePath("test_replace_nonexistent_output.docx");
        Assert.Throws<FileNotFoundException>(() =>
            _tool.Execute("replace", docPath, outputPath: outputPath, imageIndex: 0,
                newImagePath: "C:\\nonexistent\\new_image.png"));
    }

    [Fact]
    public void ExtractImages_WithInvalidIndex_ShouldThrowArgumentException()
    {
        var docPath = CreateWordDocument("test_extract_invalid_index.docx");
        var imagePath = CreateTestImage("test_extract_invalid.png");
        var doc = new Document(docPath);
        var builder = new DocumentBuilder(doc);
        builder.InsertImage(imagePath);
        doc.Save(docPath);

        var outputDir = CreateTestFilePath("extract_invalid_output");
        Directory.CreateDirectory(outputDir);
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("extract", docPath, outputDir: outputDir, extractImageIndex: 999));

        Assert.Contains("out of range", ex.Message);
    }

    [SkippableFact]
    public void ExtractImages_WithNoImages_ShouldReturnMessage()
    {
        SkipInEvaluationMode(AsposeLibraryType.Words, "Evaluation watermark is treated as an image");
        var docPath = CreateWordDocumentWithContent("test_extract_no_images.docx", "No images");
        var outputDir = CreateTestFilePath("extract_no_images_output");
        Directory.CreateDirectory(outputDir);
        var result = _tool.Execute("extract", docPath, outputDir: outputDir);
        Assert.Contains("No images found", result);
    }

    #endregion

    #region Session

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

        var sessionDoc = SessionManager.GetDocument<Document>(sessionId);
        var shapes = sessionDoc.GetChildNodes(NodeType.Shape, true).Cast<Shape>().Where(s => s.HasImage).ToList();
        Assert.True(shapes.Count > 0);
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

        var sessionDoc = SessionManager.GetDocument<Document>(sessionId);
        var shapes = sessionDoc.GetChildNodes(NodeType.Shape, true).Cast<Shape>().Where(s => s.HasImage).ToList();
        Assert.True(shapes.Count > 0);
    }

    #endregion
}