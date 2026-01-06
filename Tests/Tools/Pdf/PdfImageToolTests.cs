using System.Drawing;
using System.Drawing.Imaging;
using System.Runtime.Versioning;
using System.Text.Json;
using Aspose.Pdf;
using AsposeMcpServer.Tests.Helpers;
using AsposeMcpServer.Tools.Pdf;
using DrawingColor = System.Drawing.Color;
using Rectangle = Aspose.Pdf.Rectangle;

namespace AsposeMcpServer.Tests.Tools.Pdf;

[SupportedOSPlatform("windows")]
public class PdfImageToolTests : PdfTestBase
{
    private readonly PdfImageTool _tool;

    public PdfImageToolTests()
    {
        _tool = new PdfImageTool(SessionManager);
    }

    private string CreateTestImage(string fileName)
    {
        var imagePath = CreateTestFilePath(fileName);
        using var bitmap = new Bitmap(100, 100);
        using (var g = Graphics.FromImage(bitmap))
        {
            g.Clear(DrawingColor.Red);
        }

        bitmap.Save(imagePath, ImageFormat.Png);
        return imagePath;
    }

    private string CreateTestPdf(string fileName, int pageCount = 1)
    {
        var filePath = CreateTestFilePath(fileName);
        using var document = new Document();
        for (var i = 0; i < pageCount; i++)
            document.Pages.Add();
        document.Save(filePath);
        return filePath;
    }

    private string CreatePdfWithImage(string fileName)
    {
        var filePath = CreateTestFilePath(fileName);
        using var document = new Document();
        var page = document.Pages.Add();
        var imagePath = CreateTestImage($"img_{Guid.NewGuid()}.png");
        using (var imageStream = new FileStream(imagePath, FileMode.Open))
        {
            page.AddImage(imageStream, new Rectangle(100, 100, 300, 300));
        }

        document.Save(filePath);
        return filePath;
    }

    #region General

    [Fact]
    public void Add_ShouldAddImageToPdf()
    {
        var pdfPath = CreateTestPdf("test_add.pdf");
        var imagePath = CreateTestImage("test_add_image.png");
        var outputPath = CreateTestFilePath("test_add_output.pdf");
        var result = _tool.Execute("add", pdfPath, outputPath: outputPath,
            pageIndex: 1, imagePath: imagePath, x: 100, y: 100);
        Assert.StartsWith("Added image", result);
        using var document = new Document(outputPath);
        Assert.True(document.Pages[1].Resources.Images.Count > 0);
    }

    [Fact]
    public void Add_WithWidthHeight_ShouldAddWithSize()
    {
        var pdfPath = CreateTestPdf("test_add_size.pdf");
        var imagePath = CreateTestImage("test_add_size_image.png");
        var outputPath = CreateTestFilePath("test_add_size_output.pdf");
        var result = _tool.Execute("add", pdfPath, outputPath: outputPath,
            pageIndex: 1, imagePath: imagePath, x: 100, y: 100, width: 300, height: 300);
        Assert.StartsWith("Added image", result);
        Assert.True(File.Exists(outputPath));
    }

    [Fact]
    public void Delete_ShouldDeleteImage()
    {
        var pdfPath = CreatePdfWithImage("test_delete.pdf");
        var outputPath = CreateTestFilePath("test_delete_output.pdf");
        var result = _tool.Execute("delete", pdfPath, outputPath: outputPath,
            pageIndex: 1, imageIndex: 1);
        Assert.StartsWith("Deleted image", result);
        using var document = new Document(outputPath);
        Assert.Empty(document.Pages[1].Resources.Images);
    }

    [Fact]
    public void Edit_WithImagePath_ShouldReplaceImage()
    {
        var pdfPath = CreatePdfWithImage("test_edit_replace.pdf");
        var newImagePath = CreateTestImage("test_edit_new.png");
        var outputPath = CreateTestFilePath("test_edit_replace_output.pdf");
        var result = _tool.Execute("edit", pdfPath, outputPath: outputPath,
            pageIndex: 1, imageIndex: 1, imagePath: newImagePath, x: 200, y: 200);
        Assert.StartsWith("Replaced", result);
        Assert.True(File.Exists(outputPath));
    }

    [Fact]
    public void Edit_WithoutImagePath_ShouldMoveImage()
    {
        var pdfPath = CreatePdfWithImage("test_edit_move.pdf");
        var outputPath = CreateTestFilePath("test_edit_move_output.pdf");
        var result = _tool.Execute("edit", pdfPath, outputPath: outputPath,
            pageIndex: 1, imageIndex: 1, x: 300, y: 300);
        Assert.StartsWith("Moved", result);
        Assert.True(File.Exists(outputPath));
    }

    [Fact]
    public void Edit_WithWidthHeight_ShouldResizeImage()
    {
        var pdfPath = CreatePdfWithImage("test_edit_resize.pdf");
        var outputPath = CreateTestFilePath("test_edit_resize_output.pdf");
        var result = _tool.Execute("edit", pdfPath, outputPath: outputPath,
            pageIndex: 1, imageIndex: 1, width: 400, height: 400);
        Assert.StartsWith("Moved", result);
        Assert.True(File.Exists(outputPath));
    }

    [Fact]
    public void Extract_ShouldExtractImage()
    {
        var pdfPath = CreatePdfWithImage("test_extract.pdf");
        var outputImagePath = CreateTestFilePath("test_extracted.png");
        var result = _tool.Execute("extract", pdfPath,
            outputPath: outputImagePath, pageIndex: 1, imageIndex: 1);
        Assert.StartsWith("Extracted image", result);
        Assert.True(File.Exists(outputImagePath));
    }

    [Fact]
    public void Extract_WithOutputDir_ShouldExtractAllImages()
    {
        var pdfPath = CreatePdfWithImage("test_extract_all.pdf");
        var outputDir = Path.Combine(Path.GetTempPath(), $"PdfImageTest_{Guid.NewGuid()}");
        Directory.CreateDirectory(outputDir);
        try
        {
            var result = _tool.Execute("extract", pdfPath,
                outputDir: outputDir, pageIndex: 1);
            Assert.StartsWith("Extracted", result);
            var files = Directory.GetFiles(outputDir, "*.png");
            Assert.NotEmpty(files);
        }
        finally
        {
            if (Directory.Exists(outputDir))
                Directory.Delete(outputDir, true);
        }
    }

    [Fact]
    public void Get_WithImages_ShouldReturnImageInfo()
    {
        var pdfPath = CreatePdfWithImage("test_get.pdf");
        var result = _tool.Execute("get", pdfPath, pageIndex: 1);
        var json = JsonSerializer.Deserialize<JsonElement>(result);
        Assert.True(json.GetProperty("count").GetInt32() > 0);
        Assert.True(json.TryGetProperty("items", out _));
    }

    [Fact]
    public void Get_WithNoImages_ShouldReturnEmptyResult()
    {
        var pdfPath = CreateTestPdf("test_get_empty.pdf");
        var result = _tool.Execute("get", pdfPath, pageIndex: 1);
        var json = JsonSerializer.Deserialize<JsonElement>(result);
        Assert.Equal(0, json.GetProperty("count").GetInt32());
        Assert.Contains("No images found", result);
    }

    [Fact]
    public void Get_WithoutPageIndex_ShouldReturnAllImages()
    {
        var pdfPath = CreatePdfWithImage("test_get_all.pdf");
        var result = _tool.Execute("get", pdfPath);
        var json = JsonSerializer.Deserialize<JsonElement>(result);
        Assert.True(json.TryGetProperty("count", out _));
    }

    [Theory]
    [InlineData("ADD")]
    [InlineData("Add")]
    [InlineData("add")]
    public void Operation_ShouldBeCaseInsensitive_Add(string operation)
    {
        var pdfPath = CreateTestPdf($"test_case_{operation}.pdf");
        var imagePath = CreateTestImage($"test_case_{operation}.png");
        var outputPath = CreateTestFilePath($"test_case_{operation}_output.pdf");
        var result = _tool.Execute(operation, pdfPath, outputPath: outputPath,
            pageIndex: 1, imagePath: imagePath, x: 100, y: 100);
        Assert.StartsWith("Added", result);
    }

    [Theory]
    [InlineData("GET")]
    [InlineData("Get")]
    [InlineData("get")]
    public void Operation_ShouldBeCaseInsensitive_Get(string operation)
    {
        var pdfPath = CreateTestPdf($"test_case_get_{operation}.pdf");
        var result = _tool.Execute(operation, pdfPath, pageIndex: 1);
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
    public void Add_WithInvalidPageIndex_ShouldThrowArgumentException()
    {
        var pdfPath = CreateTestPdf("test_add_invalid_page.pdf");
        var imagePath = CreateTestImage("test_add_invalid_page.png");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("add", pdfPath, pageIndex: 99, imagePath: imagePath, x: 100, y: 100));
        Assert.Contains("pageIndex must be between", ex.Message);
    }

    [Fact]
    public void Add_WithMissingImagePath_ShouldThrowArgumentException()
    {
        var pdfPath = CreateTestPdf("test_add_no_image.pdf");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("add", pdfPath, pageIndex: 1, x: 100, y: 100));
        Assert.Contains("imagePath", ex.Message);
    }

    [Fact]
    public void Add_WithNonExistentImagePath_ShouldThrowFileNotFoundException()
    {
        var pdfPath = CreateTestPdf("test_add_nonexistent.pdf");
        Assert.Throws<FileNotFoundException>(() =>
            _tool.Execute("add", pdfPath, pageIndex: 1, imagePath: @"C:\nonexistent\image.png", x: 100, y: 100));
    }

    [Fact]
    public void Delete_WithInvalidImageIndex_ShouldThrowArgumentException()
    {
        var pdfPath = CreatePdfWithImage("test_delete_invalid.pdf");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("delete", pdfPath, pageIndex: 1, imageIndex: 99));
        Assert.Contains("imageIndex must be between", ex.Message);
    }

    [Fact]
    public void Delete_WithNoImages_ShouldThrowArgumentException()
    {
        var pdfPath = CreateTestPdf("test_delete_no_images.pdf");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("delete", pdfPath, pageIndex: 1, imageIndex: 1));
        Assert.Contains("imageIndex must be between", ex.Message);
    }

    [Fact]
    public void Edit_WithInvalidImageIndex_ShouldThrowArgumentException()
    {
        var pdfPath = CreatePdfWithImage("test_edit_invalid.pdf");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("edit", pdfPath, pageIndex: 1, imageIndex: 99, x: 100, y: 100));
        Assert.Contains("imageIndex must be between", ex.Message);
    }

    [Fact]
    public void Extract_WithInvalidPageIndex_ShouldThrowArgumentException()
    {
        var pdfPath = CreatePdfWithImage("test_extract_invalid_page.pdf");
        var outputPath = CreateTestFilePath("test_extract_invalid.png");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("extract", pdfPath, outputPath: outputPath, pageIndex: 99, imageIndex: 1));
        Assert.Contains("pageIndex must be between", ex.Message);
    }

    [Fact]
    public void Extract_WithInvalidImageIndex_ShouldThrowArgumentException()
    {
        var pdfPath = CreatePdfWithImage("test_extract_invalid_img.pdf");
        var outputPath = CreateTestFilePath("test_extract_invalid_img.png");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("extract", pdfPath, outputPath: outputPath, pageIndex: 1, imageIndex: 99));
        Assert.Contains("imageIndex must be between", ex.Message);
    }

    [Fact]
    public void Extract_WithNoImages_ShouldReturnNoImagesMessage()
    {
        var pdfPath = CreateTestPdf("test_extract_no_images.pdf");
        var outputPath = CreateTestFilePath("test_extract_no_images.png");
        var result = _tool.Execute("extract", pdfPath, outputPath: outputPath, pageIndex: 1);
        Assert.Contains("No images found", result);
    }

    [Fact]
    public void Get_WithInvalidPageIndex_ShouldThrowArgumentException()
    {
        var pdfPath = CreateTestPdf("test_get_invalid_page.pdf");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("get", pdfPath, pageIndex: 99));
        Assert.Contains("pageIndex must be between", ex.Message);
    }

    [Fact]
    public void Execute_WithNoPathOrSessionId_ShouldThrowException()
    {
        var ex = Assert.Throws<ArgumentException>(() => _tool.Execute("get"));
        Assert.Contains("path", ex.Message.ToLower());
    }

    #endregion

    #region Session

    [Fact]
    public void Get_WithSessionId_ShouldGetFromSession()
    {
        var pdfPath = CreatePdfWithImage("test_session_get.pdf");
        var sessionId = OpenSession(pdfPath);
        var result = _tool.Execute("get", sessionId: sessionId, pageIndex: 1);
        var json = JsonSerializer.Deserialize<JsonElement>(result);
        Assert.True(json.GetProperty("count").GetInt32() > 0);
    }

    [Fact]
    public void Add_WithSessionId_ShouldAddToSession()
    {
        var pdfPath = CreateTestPdf("test_session_add.pdf");
        var sessionId = OpenSession(pdfPath);
        var imagePath = CreateTestImage("test_session_add.png");
        var docBefore = SessionManager.GetDocument<Document>(sessionId);
        var countBefore = docBefore.Pages[1].Resources.Images.Count;
        var result = _tool.Execute("add", sessionId: sessionId,
            pageIndex: 1, imagePath: imagePath, x: 100, y: 100);
        Assert.StartsWith("Added image", result);
        Assert.Contains("session", result);
        var docAfter = SessionManager.GetDocument<Document>(sessionId);
        Assert.True(docAfter.Pages[1].Resources.Images.Count > countBefore);
    }

    [Fact]
    public void Delete_WithSessionId_ShouldDeleteFromSession()
    {
        var pdfPath = CreatePdfWithImage("test_session_delete.pdf");
        var sessionId = OpenSession(pdfPath);
        var docBefore = SessionManager.GetDocument<Document>(sessionId);
        var countBefore = docBefore.Pages[1].Resources.Images.Count;
        Assert.True(countBefore > 0);
        var result = _tool.Execute("delete", sessionId: sessionId,
            pageIndex: 1, imageIndex: 1);
        Assert.StartsWith("Deleted", result);
        Assert.Contains("session", result);
        var docAfter = SessionManager.GetDocument<Document>(sessionId);
        Assert.True(docAfter.Pages[1].Resources.Images.Count < countBefore);
    }

    [Fact]
    public void Edit_WithSessionId_ShouldEditInSession()
    {
        var pdfPath = CreatePdfWithImage("test_session_edit.pdf");
        var sessionId = OpenSession(pdfPath);
        var result = _tool.Execute("edit", sessionId: sessionId,
            pageIndex: 1, imageIndex: 1, x: 300, y: 300);
        Assert.StartsWith("Moved", result);
        Assert.Contains("session", result);
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
        var pdfPath1 = CreateTestPdf("test_path_image.pdf");
        var pdfPath2 = CreatePdfWithImage("test_session_image.pdf");
        var sessionId = OpenSession(pdfPath2);
        var result = _tool.Execute("get", pdfPath1, sessionId, pageIndex: 1);
        var json = JsonSerializer.Deserialize<JsonElement>(result);
        Assert.True(json.GetProperty("count").GetInt32() > 0);
    }

    #endregion
}