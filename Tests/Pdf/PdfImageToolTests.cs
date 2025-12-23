using System.Drawing;
using System.Drawing.Imaging;
using System.Text.Json.Nodes;
using Aspose.Pdf;
using AsposeMcpServer.Tests.Helpers;
using AsposeMcpServer.Tools.Pdf;
using DrawingColor = System.Drawing.Color;
using Rectangle = Aspose.Pdf.Rectangle;

namespace AsposeMcpServer.Tests.Pdf;

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
        Assert.Contains("Image", result, StringComparison.OrdinalIgnoreCase);
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
}