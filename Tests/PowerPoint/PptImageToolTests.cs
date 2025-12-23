using System.Drawing;
using System.Drawing.Imaging;
using System.Text.Json.Nodes;
using Aspose.Slides;
using Aspose.Slides.Export;
using AsposeMcpServer.Tests.Helpers;
using AsposeMcpServer.Tools.PowerPoint;

namespace AsposeMcpServer.Tests.PowerPoint;

public class PptImageToolTests : TestBase
{
    private readonly PptImageTool _tool = new();

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
        presentation.Slides.AddEmptySlide(presentation.LayoutSlides[0]);
        presentation.Save(filePath, SaveFormat.Pptx);
        return filePath;
    }

    [Fact]
    public async Task AddImage_ShouldAddImageToSlide()
    {
        // Arrange
        var pptPath = CreateTestPresentation("test_add_image.pptx");
        var imagePath = CreateTestImage("test_image.png");
        var outputPath = CreateTestFilePath("test_add_image_output.pptx");
        var arguments = new JsonObject
        {
            ["operation"] = "add",
            ["path"] = pptPath,
            ["outputPath"] = outputPath,
            ["slideIndex"] = 0,
            ["imagePath"] = imagePath,
            ["x"] = 100,
            ["y"] = 100,
            ["width"] = 200,
            ["height"] = 150
        };

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        using var presentation = new Presentation(outputPath);
        var slide = presentation.Slides[0];
        var images = slide.Shapes.OfType<IPictureFrame>().ToList();
        Assert.True(images.Count > 0, "Slide should contain at least one image");
    }

    [Fact]
    public async Task EditImage_ShouldModifyImage()
    {
        // Arrange
        var pptPath = CreateTestPresentation("test_edit_image.pptx");
        var imagePath = CreateTestImage("test_image2.png");
        using (var ppt = new Presentation(pptPath))
        {
            var pptSlide = ppt.Slides[0];
            await using var imageStream = File.OpenRead(imagePath);
            var pictureImage = ppt.Images.AddImage(imageStream);
            pptSlide.Shapes.AddPictureFrame(ShapeType.Rectangle, 100, 100, 200, 150, pictureImage);
            ppt.Save(pptPath, SaveFormat.Pptx);
        }

        var outputPath = CreateTestFilePath("test_edit_image_output.pptx");
        var arguments = new JsonObject
        {
            ["operation"] = "edit",
            ["path"] = pptPath,
            ["outputPath"] = outputPath,
            ["slideIndex"] = 0,
            ["shapeIndex"] = 0,
            ["width"] = 300,
            ["height"] = 200
        };

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        using var presentation = new Presentation(outputPath);
        var slide = presentation.Slides[0];
        // Find the image by type (using image-specific index logic)
        var images = slide.Shapes.OfType<IPictureFrame>().ToList();
        Assert.True(images.Count > 0, "Slide should contain at least one image after editing");
        // The edited image should be at index 0 (first image)
        var image = images[0];
        Assert.NotNull(image);
        Assert.Equal(300, image.Width);
        Assert.Equal(200, image.Height);
    }
}