using System.Drawing;
using System.Drawing.Imaging;
using System.Runtime.Versioning;
using System.Text.Json.Nodes;
using Aspose.Pdf;
using AsposeMcpServer.Tests.Helpers;
using AsposeMcpServer.Tools.Pdf;
using DrawingColor = System.Drawing.Color;
using Rectangle = Aspose.Pdf.Rectangle;

namespace AsposeMcpServer.Tests.Pdf;

[SupportedOSPlatform("windows")]
public class PdfImageToolTests : PdfTestBase
{
    private readonly PdfImageTool _tool = new();

    private string CreateTestImage(string fileName)
    {
        var imagePath = CreateTestFilePath(fileName);
        using var bitmap = new Bitmap(1, 1);
        bitmap.SetPixel(0, 0, DrawingColor.Red);
        bitmap.Save(imagePath, ImageFormat.Png);
        return imagePath;
    }

    [Fact]
    public async Task AddImage_ShouldAddImageToPdf()
    {
        // Arrange
        var pdfPath = CreateTestFilePath("test_add_image.pdf");
        var document = new Document();
        document.Pages.Add();
        document.Save(pdfPath);

        var imagePath = CreateTestImage("test_image.png");
        var outputPath = CreateTestFilePath("test_add_image_output.pdf");
        var arguments = new JsonObject
        {
            ["operation"] = "add",
            ["path"] = pdfPath,
            ["outputPath"] = outputPath,
            ["imagePath"] = imagePath,
            ["pageIndex"] = 0,
            ["x"] = 100,
            ["y"] = 100
        };

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.True(File.Exists(outputPath), "PDF with image should be created");
        var doc = new Document(outputPath);
        Assert.True(doc.Pages.Count > 0, "PDF should have pages");
    }

    [Fact]
    public async Task GetImages_ShouldReturnAllImages()
    {
        // Arrange
        var pdfPath = CreateTestFilePath("test_get_images.pdf");
        var document = new Document();
        var page = document.Pages.Add();
        var imagePath = CreateTestImage("test_image2.png");
        await using (var imageStream = new FileStream(imagePath, FileMode.Open))
        {
            var rect = new Rectangle(100, 100, 300, 300);
            page.AddImage(imageStream, rect);
        }

        document.Save(pdfPath);

        var arguments = new JsonObject
        {
            ["operation"] = "get",
            ["path"] = pdfPath
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.NotNull(result);
        Assert.NotEmpty(result);
        Assert.Contains("count", result);
    }

    [Fact]
    public async Task DeleteImage_ShouldDeleteImage()
    {
        // Arrange
        var pdfPath = CreateTestFilePath("test_delete_image.pdf");
        var document = new Document();
        var page = document.Pages.Add();
        var imagePath = CreateTestImage("test_image3.png");
        await using (var imageStream = new FileStream(imagePath, FileMode.Open))
        {
            var rect = new Rectangle(100, 100, 300, 300);
            page.AddImage(imageStream, rect);
        }

        document.Save(pdfPath);

        var outputPath = CreateTestFilePath("test_delete_image_output.pdf");
        var arguments = new JsonObject
        {
            ["operation"] = "delete",
            ["path"] = pdfPath,
            ["outputPath"] = outputPath,
            ["imageIndex"] = 0,
            ["pageIndex"] = 0
        };

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.True(File.Exists(outputPath), "Output PDF should be created");
    }

    [Fact]
    public async Task EditImage_ShouldEditImagePosition()
    {
        // Arrange
        var pdfPath = CreateTestFilePath("test_edit_image.pdf");
        var document = new Document();
        var page = document.Pages.Add();
        var originalImagePath = CreateTestImage("test_image4.png");
        await using (var imageStream = new FileStream(originalImagePath, FileMode.Open))
        {
            var rect = new Rectangle(100, 100, 300, 300);
            page.AddImage(imageStream, rect);
        }

        document.Save(pdfPath);

        var imagePath = CreateTestImage("test_image6.png");
        var outputPath = CreateTestFilePath("test_edit_image_output.pdf");
        var arguments = new JsonObject
        {
            ["operation"] = "edit",
            ["path"] = pdfPath,
            ["outputPath"] = outputPath,
            ["imageIndex"] = 1,
            ["pageIndex"] = 1,
            ["imagePath"] = imagePath,
            ["x"] = 200,
            ["y"] = 200
        };

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.True(File.Exists(outputPath), "Output PDF should be created");
    }

    [Fact]
    public async Task ExtractImage_ShouldExtractImage()
    {
        // Arrange
        var pdfPath = CreateTestFilePath("test_extract_image.pdf");
        var document = new Document();
        var page = document.Pages.Add();
        var imagePath = CreateTestImage("test_image5.png");
        await using (var imageStream = new FileStream(imagePath, FileMode.Open))
        {
            var rect = new Rectangle(100, 100, 300, 300);
            page.AddImage(imageStream, rect);
        }

        document.Save(pdfPath);

        var outputImagePath = CreateTestFilePath("test_extracted_image.png");
        var arguments = new JsonObject
        {
            ["operation"] = "extract",
            ["path"] = pdfPath,
            ["imageIndex"] = 1,
            ["pageIndex"] = 1,
            ["outputPath"] = outputImagePath
        };

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.True(File.Exists(outputImagePath), "Extracted image should be created");
    }

    [Fact]
    public async Task AddImage_WithJpegFormat_ShouldAddJpeg()
    {
        // Arrange
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
        var arguments = new JsonObject
        {
            ["operation"] = "add",
            ["path"] = pdfPath,
            ["outputPath"] = outputPath,
            ["imagePath"] = imagePath,
            ["pageIndex"] = 0,
            ["x"] = 50,
            ["y"] = 50
        };

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.True(File.Exists(outputPath), "PDF with JPEG image should be created");
    }

    [Fact]
    public async Task EditImage_WithWidthHeight_ShouldResizeImage()
    {
        // Arrange
        var pdfPath = CreateTestFilePath("test_resize_image.pdf");
        var document = new Document();
        var page = document.Pages.Add();
        var imagePath = CreateTestImage("test_resize_source.png");
        await using (var imageStream = new FileStream(imagePath, FileMode.Open))
        {
            var rect = new Rectangle(100, 100, 200, 200);
            page.AddImage(imageStream, rect);
        }

        document.Save(pdfPath);

        var newImagePath = CreateTestImage("test_resize_new.png");
        var outputPath = CreateTestFilePath("test_resize_image_output.pdf");
        var arguments = new JsonObject
        {
            ["operation"] = "edit",
            ["path"] = pdfPath,
            ["outputPath"] = outputPath,
            ["imageIndex"] = 1,
            ["pageIndex"] = 1,
            ["imagePath"] = newImagePath,
            ["width"] = 300,
            ["height"] = 300
        };

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.True(File.Exists(outputPath), "PDF with resized image should be created");
    }

    [Fact]
    public async Task Extract_WithOutputDir_ShouldExtractToDirectory()
    {
        // Arrange
        var pdfPath = CreateTestFilePath("test_extract_to_dir.pdf");
        var document = new Document();
        var page = document.Pages.Add();
        var imagePath = CreateTestImage("test_dir_source.png");
        await using (var imageStream = new FileStream(imagePath, FileMode.Open))
        {
            var rect = new Rectangle(100, 100, 300, 300);
            page.AddImage(imageStream, rect);
        }

        document.Save(pdfPath);

        var outputDir = Path.Combine(Path.GetTempPath(), $"PdfImageTest_{Guid.NewGuid()}");
        Directory.CreateDirectory(outputDir);

        var arguments = new JsonObject
        {
            ["operation"] = "extract",
            ["path"] = pdfPath,
            ["imageIndex"] = 1,
            ["pageIndex"] = 1,
            ["outputPath"] = Path.Combine(outputDir, "extracted.png")
        };

        try
        {
            // Act
            await _tool.ExecuteAsync(arguments);

            // Assert
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
    public async Task Get_FromMultiplePages_ShouldGetAllImages()
    {
        // Arrange
        var pdfPath = CreateTestFilePath("test_multi_page_images.pdf");
        var document = new Document();

        // Add images to multiple pages
        for (var i = 0; i < 3; i++)
        {
            var page = document.Pages.Add();
            var imagePath = CreateTestImage($"test_multi_page_img_{i}.png");
            await using var imageStream = new FileStream(imagePath, FileMode.Open);
            var rect = new Rectangle(100, 100, 200, 200);
            page.AddImage(imageStream, rect);
        }

        document.Save(pdfPath);

        var arguments = new JsonObject
        {
            ["operation"] = "get",
            ["path"] = pdfPath
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.NotNull(result);
        Assert.Contains("count", result);
    }

    [Fact]
    public async Task EditImage_WithoutImagePath_ShouldMoveExistingImage()
    {
        // Arrange
        var pdfPath = CreateTestFilePath("test_move_image.pdf");
        var document = new Document();
        var page = document.Pages.Add();
        var imagePath = CreateTestImage("test_move_source.png");
        await using (var imageStream = new FileStream(imagePath, FileMode.Open))
        {
            var rect = new Rectangle(100, 100, 200, 200);
            page.AddImage(imageStream, rect);
        }

        document.Save(pdfPath);

        var outputPath = CreateTestFilePath("test_move_image_output.pdf");
        var arguments = new JsonObject
        {
            ["operation"] = "edit",
            ["path"] = pdfPath,
            ["outputPath"] = outputPath,
            ["imageIndex"] = 1,
            ["pageIndex"] = 1,
            ["x"] = 300,
            ["y"] = 300
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.True(File.Exists(outputPath), "Output PDF should be created");
        Assert.Contains("Moved", result);
    }

    [Fact]
    public async Task AddImage_WithInvalidPageIndex_ShouldThrowArgumentException()
    {
        // Arrange
        var pdfPath = CreateTestFilePath("test_add_invalid_page.pdf");
        var document = new Document();
        document.Pages.Add();
        document.Save(pdfPath);

        var imagePath = CreateTestImage("test_invalid_page_image.png");
        var arguments = new JsonObject
        {
            ["operation"] = "add",
            ["path"] = pdfPath,
            ["imagePath"] = imagePath,
            ["pageIndex"] = 99,
            ["x"] = 100,
            ["y"] = 100
        };

        // Act & Assert
        var exception = await Assert.ThrowsAsync<ArgumentException>(() => _tool.ExecuteAsync(arguments));
        Assert.Contains("pageIndex must be between", exception.Message);
    }

    [Fact]
    public async Task DeleteImage_WithInvalidImageIndex_ShouldThrowArgumentException()
    {
        // Arrange
        var pdfPath = CreateTestFilePath("test_delete_invalid_index.pdf");
        var document = new Document();
        var page = document.Pages.Add();
        var imagePath = CreateTestImage("test_delete_invalid_image.png");
        await using (var imageStream = new FileStream(imagePath, FileMode.Open))
        {
            var rect = new Rectangle(100, 100, 200, 200);
            page.AddImage(imageStream, rect);
        }

        document.Save(pdfPath);

        var arguments = new JsonObject
        {
            ["operation"] = "delete",
            ["path"] = pdfPath,
            ["imageIndex"] = 99,
            ["pageIndex"] = 1
        };

        // Act & Assert
        var exception = await Assert.ThrowsAsync<ArgumentException>(() => _tool.ExecuteAsync(arguments));
        Assert.Contains("imageIndex must be between", exception.Message);
    }

    [Fact]
    public async Task Execute_WithUnknownOperation_ShouldThrowArgumentException()
    {
        // Arrange
        var pdfPath = CreateTestFilePath("test_unknown_op.pdf");
        var document = new Document();
        document.Pages.Add();
        document.Save(pdfPath);

        var arguments = new JsonObject
        {
            ["operation"] = "unknown",
            ["path"] = pdfPath
        };

        // Act & Assert
        var exception = await Assert.ThrowsAsync<ArgumentException>(() => _tool.ExecuteAsync(arguments));
        Assert.Contains("Unknown operation", exception.Message);
    }

    [Fact]
    public async Task GetImages_WithNoImages_ShouldReturnEmptyResult()
    {
        // Arrange
        var pdfPath = CreateTestFilePath("test_get_no_images.pdf");
        var document = new Document();
        document.Pages.Add();
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
        Assert.Contains("\"count\": 0", result);
        Assert.Contains("No images found", result);
    }

    [Fact]
    public async Task AddImage_WithNonExistentFile_ShouldThrowFileNotFoundException()
    {
        // Arrange
        var pdfPath = CreateTestFilePath("test_add_nonexistent.pdf");
        var document = new Document();
        document.Pages.Add();
        document.Save(pdfPath);

        var arguments = new JsonObject
        {
            ["operation"] = "add",
            ["path"] = pdfPath,
            ["imagePath"] = @"C:\nonexistent\image.png",
            ["pageIndex"] = 1,
            ["x"] = 100,
            ["y"] = 100
        };

        // Act & Assert
        await Assert.ThrowsAsync<FileNotFoundException>(() => _tool.ExecuteAsync(arguments));
    }
}