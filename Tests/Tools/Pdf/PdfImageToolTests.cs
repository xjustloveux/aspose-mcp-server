using System.Drawing;
using System.Drawing.Imaging;
using System.Runtime.Versioning;
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
        using var bitmap = new Bitmap(1, 1);
        bitmap.SetPixel(0, 0, DrawingColor.Red);
        bitmap.Save(imagePath, ImageFormat.Png);
        return imagePath;
    }

    private string CreateTestPdf(string fileName)
    {
        var filePath = CreateTestFilePath(fileName);
        var document = new Document();
        document.Pages.Add();
        document.Save(filePath);
        return filePath;
    }

    #region General Tests

    [Fact]
    public void AddImage_ShouldAddImageToPdf()
    {
        var pdfPath = CreateTestFilePath("test_add_image.pdf");
        var document = new Document();
        document.Pages.Add();
        document.Save(pdfPath);

        var imagePath = CreateTestImage("test_image.png");
        var outputPath = CreateTestFilePath("test_add_image_output.pdf");
        _tool.Execute(
            "add",
            pdfPath,
            pageIndex: 0,
            outputPath: outputPath,
            imagePath: imagePath,
            x: 100,
            y: 100);
        Assert.True(File.Exists(outputPath), "PDF with image should be created");
        var doc = new Document(outputPath);
        Assert.True(doc.Pages.Count > 0, "PDF should have pages");
    }

    [Fact]
    public void GetImages_ShouldReturnAllImages()
    {
        var pdfPath = CreateTestFilePath("test_get_images.pdf");
        var document = new Document();
        var page = document.Pages.Add();
        var imagePath = CreateTestImage("test_image2.png");
        using (var imageStream = new FileStream(imagePath, FileMode.Open))
        {
            var rect = new Rectangle(100, 100, 300, 300);
            page.AddImage(imageStream, rect);
        }

        document.Save(pdfPath);
        var result = _tool.Execute("get", pdfPath);
        Assert.NotNull(result);
        Assert.NotEmpty(result);
        Assert.Contains("count", result);
    }

    [Fact]
    public void DeleteImage_ShouldDeleteImage()
    {
        var pdfPath = CreateTestFilePath("test_delete_image.pdf");
        var document = new Document();
        var page = document.Pages.Add();
        var imagePath = CreateTestImage("test_image3.png");
        using (var imageStream = new FileStream(imagePath, FileMode.Open))
        {
            var rect = new Rectangle(100, 100, 300, 300);
            page.AddImage(imageStream, rect);
        }

        document.Save(pdfPath);

        var outputPath = CreateTestFilePath("test_delete_image_output.pdf");
        _tool.Execute(
            "delete",
            pdfPath,
            pageIndex: 0,
            outputPath: outputPath,
            imageIndex: 0);
        Assert.True(File.Exists(outputPath), "Output PDF should be created");
    }

    [Fact]
    public void EditImage_ShouldEditImagePosition()
    {
        var pdfPath = CreateTestFilePath("test_edit_image.pdf");
        var document = new Document();
        var page = document.Pages.Add();
        var originalImagePath = CreateTestImage("test_image4.png");
        using (var imageStream = new FileStream(originalImagePath, FileMode.Open))
        {
            var rect = new Rectangle(100, 100, 300, 300);
            page.AddImage(imageStream, rect);
        }

        document.Save(pdfPath);

        var imagePath = CreateTestImage("test_image6.png");
        var outputPath = CreateTestFilePath("test_edit_image_output.pdf");
        _tool.Execute(
            "edit",
            pdfPath,
            pageIndex: 1,
            outputPath: outputPath,
            imagePath: imagePath,
            imageIndex: 1,
            x: 200,
            y: 200);
        Assert.True(File.Exists(outputPath), "Output PDF should be created");
    }

    [Fact]
    public void ExtractImage_ShouldExtractImage()
    {
        var pdfPath = CreateTestFilePath("test_extract_image.pdf");
        var document = new Document();
        var page = document.Pages.Add();
        var imagePath = CreateTestImage("test_image5.png");
        using (var imageStream = new FileStream(imagePath, FileMode.Open))
        {
            var rect = new Rectangle(100, 100, 300, 300);
            page.AddImage(imageStream, rect);
        }

        document.Save(pdfPath);

        var outputImagePath = CreateTestFilePath("test_extracted_image.png");
        _tool.Execute(
            "extract",
            pdfPath,
            pageIndex: 1,
            outputPath: outputImagePath,
            imageIndex: 1);
        Assert.True(File.Exists(outputImagePath), "Extracted image should be created");
    }

    [Fact]
    public void AddImage_WithJpegFormat_ShouldAddJpeg()
    {
        var pdfPath = CreateTestFilePath("test_add_jpeg.pdf");
        var document = new Document();
        document.Pages.Add();
        document.Save(pdfPath);

        // Create JPEG image
        var imagePath = CreateTestFilePath("test_jpeg_image.jpg");
        using (var bitmap = new Bitmap(100, 100))
        {
            using (var g = Graphics.FromImage(bitmap))
            {
                g.Clear(DrawingColor.Blue);
            }

            bitmap.Save(imagePath, ImageFormat.Jpeg);
        }

        var outputPath = CreateTestFilePath("test_add_jpeg_output.pdf");
        _tool.Execute(
            "add",
            pdfPath,
            pageIndex: 0,
            outputPath: outputPath,
            imagePath: imagePath,
            x: 50,
            y: 50);
        Assert.True(File.Exists(outputPath), "PDF with JPEG image should be created");
    }

    [Fact]
    public void EditImage_WithWidthHeight_ShouldResizeImage()
    {
        var pdfPath = CreateTestFilePath("test_resize_image.pdf");
        var document = new Document();
        var page = document.Pages.Add();
        var imagePath = CreateTestImage("test_resize_source.png");
        using (var imageStream = new FileStream(imagePath, FileMode.Open))
        {
            var rect = new Rectangle(100, 100, 200, 200);
            page.AddImage(imageStream, rect);
        }

        document.Save(pdfPath);

        var newImagePath = CreateTestImage("test_resize_new.png");
        var outputPath = CreateTestFilePath("test_resize_image_output.pdf");
        _tool.Execute(
            "edit",
            pdfPath,
            pageIndex: 1,
            outputPath: outputPath,
            imagePath: newImagePath,
            imageIndex: 1,
            width: 300,
            height: 300);
        Assert.True(File.Exists(outputPath), "PDF with resized image should be created");
    }

    [Fact]
    public void Extract_WithOutputDir_ShouldExtractToDirectory()
    {
        var pdfPath = CreateTestFilePath("test_extract_to_dir.pdf");
        var document = new Document();
        var page = document.Pages.Add();
        var imagePath = CreateTestImage("test_dir_source.png");
        using (var imageStream = new FileStream(imagePath, FileMode.Open))
        {
            var rect = new Rectangle(100, 100, 300, 300);
            page.AddImage(imageStream, rect);
        }

        document.Save(pdfPath);

        var outputDir = Path.Combine(Path.GetTempPath(), $"PdfImageTest_{Guid.NewGuid()}");
        Directory.CreateDirectory(outputDir);

        try
        {
            _tool.Execute(
                "extract",
                pdfPath,
                pageIndex: 1,
                outputPath: Path.Combine(outputDir, "extracted.png"),
                imageIndex: 1);
            var files = Directory.GetFiles(outputDir, "*.png");
            Assert.True(files.Length > 0 || File.Exists(Path.Combine(outputDir, "extracted.png")),
                "Should extract image to directory");
        }
        finally
        {
            // Cleanup
            if (Directory.Exists(outputDir))
                Directory.Delete(outputDir, true);
        }
    }

    [Fact]
    public void Get_FromMultiplePages_ShouldGetAllImages()
    {
        var pdfPath = CreateTestFilePath("test_multi_page_images.pdf");
        var document = new Document();

        // Add images to multiple pages
        for (var i = 0; i < 3; i++)
        {
            var page = document.Pages.Add();
            var imagePath = CreateTestImage($"test_multi_page_img_{i}.png");
            using var imageStream = new FileStream(imagePath, FileMode.Open);
            var rect = new Rectangle(100, 100, 200, 200);
            page.AddImage(imageStream, rect);
        }

        document.Save(pdfPath);
        var result = _tool.Execute("get", pdfPath);
        Assert.NotNull(result);
        Assert.Contains("count", result);
    }

    [Fact]
    public void EditImage_WithoutImagePath_ShouldMoveExistingImage()
    {
        var pdfPath = CreateTestFilePath("test_move_image.pdf");
        var document = new Document();
        var page = document.Pages.Add();
        var imagePath = CreateTestImage("test_move_source.png");
        using (var imageStream = new FileStream(imagePath, FileMode.Open))
        {
            var rect = new Rectangle(100, 100, 200, 200);
            page.AddImage(imageStream, rect);
        }

        document.Save(pdfPath);

        var outputPath = CreateTestFilePath("test_move_image_output.pdf");
        var result = _tool.Execute(
            "edit",
            pdfPath,
            pageIndex: 1,
            outputPath: outputPath,
            imageIndex: 1,
            x: 300,
            y: 300);
        Assert.True(File.Exists(outputPath), "Output PDF should be created");
        Assert.Contains("Moved", result);
    }

    [Fact]
    public void AddImage_WithInvalidPageIndex_ShouldThrowArgumentException()
    {
        var pdfPath = CreateTestFilePath("test_add_invalid_page.pdf");
        var document = new Document();
        document.Pages.Add();
        document.Save(pdfPath);

        var imagePath = CreateTestImage("test_invalid_page_image.png");
        var exception = Assert.Throws<ArgumentException>(() => _tool.Execute(
            "add",
            pdfPath,
            pageIndex: 99,
            imagePath: imagePath,
            x: 100,
            y: 100));
        Assert.Contains("pageIndex must be between", exception.Message);
    }

    [Fact]
    public void DeleteImage_WithInvalidImageIndex_ShouldThrowArgumentException()
    {
        var pdfPath = CreateTestFilePath("test_delete_invalid_index.pdf");
        var document = new Document();
        var page = document.Pages.Add();
        var imagePath = CreateTestImage("test_delete_invalid_image.png");
        using (var imageStream = new FileStream(imagePath, FileMode.Open))
        {
            var rect = new Rectangle(100, 100, 200, 200);
            page.AddImage(imageStream, rect);
        }

        document.Save(pdfPath);
        var exception = Assert.Throws<ArgumentException>(() => _tool.Execute(
            "delete",
            pdfPath,
            pageIndex: 1,
            imageIndex: 99));
        Assert.Contains("imageIndex must be between", exception.Message);
    }

    [Fact]
    public void Execute_WithUnknownOperation_ShouldThrowArgumentException()
    {
        var pdfPath = CreateTestFilePath("test_unknown_op.pdf");
        var document = new Document();
        document.Pages.Add();
        document.Save(pdfPath);
        var exception = Assert.Throws<ArgumentException>(() => _tool.Execute("unknown", pdfPath));
        Assert.Contains("Unknown operation", exception.Message);
    }

    [Fact]
    public void GetImages_WithNoImages_ShouldReturnEmptyResult()
    {
        var pdfPath = CreateTestFilePath("test_get_no_images.pdf");
        var document = new Document();
        document.Pages.Add();
        document.Save(pdfPath);
        var result = _tool.Execute("get", pdfPath, pageIndex: 1);
        Assert.Contains("\"count\": 0", result);
        Assert.Contains("No images found", result);
    }

    [Fact]
    public void AddImage_WithNonExistentFile_ShouldThrowFileNotFoundException()
    {
        var pdfPath = CreateTestFilePath("test_add_nonexistent.pdf");
        var document = new Document();
        document.Pages.Add();
        document.Save(pdfPath);
        Assert.Throws<FileNotFoundException>(() => _tool.Execute(
            "add",
            pdfPath,
            pageIndex: 1,
            imagePath: @"C:\nonexistent\image.png",
            x: 100,
            y: 100));
    }

    #endregion

    #region Exception Tests

    [Fact]
    public void Execute_WithMissingRequiredPath_ShouldThrowArgumentException()
    {
        var exception = Assert.Throws<ArgumentException>(() => _tool.Execute("get"));
        Assert.Contains("path", exception.Message.ToLower());
    }

    [Fact]
    public void AddImage_WithMissingImagePath_ShouldThrowArgumentException()
    {
        var pdfPath = CreateTestPdf("test_add_missing_image.pdf");
        var exception = Assert.Throws<ArgumentException>(() => _tool.Execute(
            "add",
            pdfPath,
            pageIndex: 1,
            x: 100,
            y: 100));
        Assert.Contains("imagePath", exception.Message);
    }

    #endregion

    #region Session ID Tests

    [Fact]
    public void GetImages_WithSessionId_ShouldGetFromSession()
    {
        var pdfPath = CreateTestFilePath("test_session_get_images.pdf");
        var document = new Document();
        var page = document.Pages.Add();
        var imagePath = CreateTestImage("test_session_image.png");
        using (var imageStream = new FileStream(imagePath, FileMode.Open))
        {
            var rect = new Rectangle(100, 100, 300, 300);
            page.AddImage(imageStream, rect);
        }

        document.Save(pdfPath);

        var sessionId = OpenSession(pdfPath);
        var result = _tool.Execute("get", sessionId: sessionId, pageIndex: 1);
        Assert.NotNull(result);
        Assert.Contains("count", result);
    }

    [Fact]
    public void AddImage_WithSessionId_ShouldAddToSession()
    {
        var pdfPath = CreateTestPdf("test_session_add_image.pdf");
        var sessionId = OpenSession(pdfPath);
        var imagePath = CreateTestImage("test_session_add_image.png");
        var result = _tool.Execute(
            "add",
            sessionId: sessionId,
            pageIndex: 1,
            imagePath: imagePath,
            x: 100,
            y: 100);
        Assert.Contains("Added image", result);
        Assert.Contains("session", result);

        // Verify in-memory changes
        var doc = SessionManager.GetDocument<Document>(sessionId);
        Assert.NotNull(doc);
        Assert.True(doc.Pages[1].Resources.Images.Count > 0, "Image should be added to session document");
    }

    [Fact]
    public void DeleteImage_WithSessionId_ShouldDeleteFromSession()
    {
        var pdfPath = CreateTestFilePath("test_session_delete_image.pdf");
        var document = new Document();
        var page = document.Pages.Add();
        var imagePath = CreateTestImage("test_session_delete_img.png");
        using (var imageStream = new FileStream(imagePath, FileMode.Open))
        {
            var rect = new Rectangle(100, 100, 300, 300);
            page.AddImage(imageStream, rect);
        }

        document.Save(pdfPath);

        var sessionId = OpenSession(pdfPath);

        // Verify image exists before deletion
        var docBefore = SessionManager.GetDocument<Document>(sessionId);
        var imageCountBefore = docBefore.Pages[1].Resources.Images.Count;
        Assert.True(imageCountBefore > 0, "Image should exist before deletion");
        var result = _tool.Execute(
            "delete",
            sessionId: sessionId,
            pageIndex: 1,
            imageIndex: 1);
        Assert.Contains("Deleted", result);

        // Verify in-memory changes
        var docAfter = SessionManager.GetDocument<Document>(sessionId);
        Assert.True(docAfter.Pages[1].Resources.Images.Count < imageCountBefore,
            "Image should be deleted from session document");
    }

    #endregion
}