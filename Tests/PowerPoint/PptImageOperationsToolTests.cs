using System.Drawing;
using System.Drawing.Imaging;
using System.Runtime.Versioning;
using System.Text.Json.Nodes;
using Aspose.Slides;
using Aspose.Slides.Export;
using AsposeMcpServer.Tests.Helpers;
using AsposeMcpServer.Tools.PowerPoint;

namespace AsposeMcpServer.Tests.PowerPoint;

[SupportedOSPlatform("windows")]
public class PptImageOperationsToolTests : TestBase
{
    private readonly PptImageOperationsTool _tool = new();

    private string CreateTestImage(string fileName)
    {
        var imagePath = CreateTestFilePath(fileName);
        using var bitmap = new Bitmap(1, 1);
        bitmap.SetPixel(0, 0, Color.Red);
        bitmap.Save(imagePath, ImageFormat.Png);
        return imagePath;
    }

    private string CreateTestPresentation(string fileName)
    {
        var filePath = CreateTestFilePath(fileName);
        using var presentation = new Presentation();
        var slide = presentation.Slides.AddEmptySlide(presentation.LayoutSlides[0]);
        var imagePath = CreateTestImage("test_image.png");
        using var imageStream = File.OpenRead(imagePath);
        var pictureImage = presentation.Images.AddImage(imageStream);
        slide.Shapes.AddPictureFrame(ShapeType.Rectangle, 100, 100, 200, 150, pictureImage);
        presentation.Save(filePath, SaveFormat.Pptx);
        return filePath;
    }

    [Fact]
    public async Task ExtractImages_ShouldExtractImages()
    {
        // Arrange
        var pptPath = CreateTestPresentation("test_extract_images.pptx");
        var outputDir = Path.Combine(TestDir, "extracted_images");
        Directory.CreateDirectory(outputDir);
        var arguments = new JsonObject
        {
            ["operation"] = "extract_images",
            ["path"] = pptPath,
            ["outputDir"] = outputDir
        };

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        var files = Directory.GetFiles(outputDir);
        Assert.True(files.Length > 0, "Should extract at least one image");
    }
}