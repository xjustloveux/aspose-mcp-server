using System.Drawing;
using System.Drawing.Imaging;
using System.Runtime.Versioning;
using Aspose.Slides;
using Aspose.Slides.Export;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Results.PowerPoint.Image;
using AsposeMcpServer.Tests.Infrastructure;
using AsposeMcpServer.Tools.PowerPoint;

namespace AsposeMcpServer.Tests.Tools.PowerPoint;

/// <summary>
///     Integration tests for PptImageTool.
///     Focuses on session management, file I/O, and operation routing.
///     Detailed parameter validation and business logic tests are in Handler tests.
/// </summary>
[SupportedOSPlatform("windows")]
public class PptImageToolTests : PptTestBase
{
    private readonly PptImageTool _tool;

    public PptImageToolTests()
    {
        _tool = new PptImageTool(SessionManager);
    }

    private string CreateTestImage(string fileName, int width = 10, int height = 10, Color? color = null)
    {
        var imagePath = CreateTestFilePath(fileName);
        using var bitmap = new Bitmap(width, height);
        var fillColor = color ?? Color.Red;
        for (var x = 0; x < width; x++)
        for (var y = 0; y < height; y++)
            bitmap.SetPixel(x, y, fillColor);
        bitmap.Save(imagePath, ImageFormat.Png);
        return imagePath;
    }

    private string CreateTestPresentation(string fileName, int slideCount = 1, bool addImages = false)
    {
        var filePath = CreateTestFilePath(fileName);
        using var presentation = new Presentation();
        for (var i = 0; i < slideCount; i++)
        {
            var slide = i == 0
                ? presentation.Slides[0]
                : presentation.Slides.AddEmptySlide(presentation.LayoutSlides[0]);
            if (addImages)
            {
                var imagePath = CreateTestImage($"pres_image_{fileName}_{i}.png");
                using var imageStream = File.OpenRead(imagePath);
                var pictureImage = presentation.Images.AddImage(imageStream);
                slide.Shapes.AddPictureFrame(ShapeType.Rectangle, 100, 100, 200, 150, pictureImage);
            }
        }

        presentation.Save(filePath, SaveFormat.Pptx);
        return filePath;
    }

    #region File I/O Smoke Tests

    [Fact]
    public void Add_ShouldAddImageAndPersistToFile()
    {
        var pptPath = CreateTestPresentation("test_add_image.pptx");
        var imagePath = CreateTestImage("test_image.png");
        var outputPath = CreateTestFilePath("test_add_image_output.pptx");
        var result = _tool.Execute("add", pptPath, slideIndex: 0, imagePath: imagePath, x: 100, y: 100, width: 200,
            height: 150, outputPath: outputPath);
        var data = GetResultData<SuccessResult>(result);
        Assert.StartsWith("Image added to slide", data.Message);
        using var presentation = new Presentation(outputPath);
        Assert.NotEmpty(presentation.Slides[0].Shapes.OfType<IPictureFrame>());
    }

    [Fact]
    public void Edit_ShouldModifyImageAndPersistToFile()
    {
        var pptPath = CreateTestPresentation("test_edit_image.pptx", addImages: true);
        var outputPath = CreateTestFilePath("test_edit_image_output.pptx");
        var result = _tool.Execute("edit", pptPath, slideIndex: 0, imageIndex: 0, width: 300, height: 200,
            outputPath: outputPath);
        var data = GetResultData<SuccessResult>(result);
        Assert.StartsWith("Image", data.Message);
        Assert.Contains("updated", data.Message);
        using var presentation = new Presentation(outputPath);
        var images = presentation.Slides[0].Shapes.OfType<IPictureFrame>().ToList();
        Assert.Equal(300, images[0].Width);
    }

    [Fact]
    public void Delete_ShouldRemoveImageAndPersistToFile()
    {
        var pptPath = CreateTestPresentation("test_delete.pptx", addImages: true);
        var outputPath = CreateTestFilePath("test_delete_output.pptx");
        var result = _tool.Execute("delete", pptPath, slideIndex: 0, imageIndex: 0, outputPath: outputPath);
        var data = GetResultData<SuccessResult>(result);
        Assert.StartsWith("Image", data.Message);
        Assert.Contains("deleted from slide", data.Message);
        using var presentation = new Presentation(outputPath);
        Assert.Empty(presentation.Slides[0].Shapes.OfType<IPictureFrame>());
    }

    [Fact]
    public void Get_ShouldReturnImageInfoFromFile()
    {
        var pptPath = CreateTestPresentation("test_get.pptx", addImages: true);
        var result = _tool.Execute("get", pptPath, slideIndex: 0);
        var data = GetResultData<GetImagesPptResult>(result);
        Assert.Equal(1, data.ImageCount);
    }

    [Fact]
    public void ExportSlides_ShouldExportAllSlidesToFiles()
    {
        var pptPath = CreateTestPresentation("test_export.pptx", 3, true);
        var outputDir = Path.Combine(TestDir, "exported_slides");
        Directory.CreateDirectory(outputDir);
        var result = _tool.Execute("export_slides", pptPath, outputDir: outputDir, format: "png");
        var files = Directory.GetFiles(outputDir, "*.png");
        Assert.Equal(3, files.Length);
        var data = GetResultData<SuccessResult>(result);
        Assert.StartsWith("Exported 3 slides", data.Message);
    }

    [Fact]
    public void Extract_ShouldExtractImagesToFiles()
    {
        var pptPath = CreateTestPresentation("test_extract.pptx", 2, true);
        var outputDir = Path.Combine(TestDir, "extracted_images");
        Directory.CreateDirectory(outputDir);
        var result = _tool.Execute("extract", pptPath, outputDir: outputDir);
        var files = Directory.GetFiles(outputDir);
        Assert.NotEmpty(files);
        var data = GetResultData<SuccessResult>(result);
        Assert.StartsWith("Extracted", data.Message);
    }

    #endregion

    #region Operation Routing

    [Theory]
    [InlineData("ADD")]
    [InlineData("Add")]
    [InlineData("add")]
    public void Operation_ShouldBeCaseInsensitive(string operation)
    {
        var pptPath = CreateTestPresentation($"test_case_{operation}.pptx");
        var imagePath = CreateTestImage($"test_case_{operation}.png");
        var outputPath = CreateTestFilePath($"test_case_{operation}_output.pptx");
        var result = _tool.Execute(operation, pptPath, slideIndex: 0, imagePath: imagePath, outputPath: outputPath);
        var data = GetResultData<SuccessResult>(result);
        Assert.StartsWith("Image added to slide", data.Message);
    }

    [Fact]
    public void Execute_WithUnknownOperation_ShouldThrowArgumentException()
    {
        var pptPath = CreateTestPresentation("test_unknown_op.pptx");
        var ex = Assert.Throws<ArgumentException>(() => _tool.Execute("unknown", pptPath));
        Assert.Contains("Unknown operation", ex.Message);
    }

    [Fact]
    public void Execute_WithNoPathOrSessionId_ShouldThrowException()
    {
        Assert.ThrowsAny<Exception>(() => _tool.Execute("get", slideIndex: 0));
    }

    #endregion

    #region Session Management

    [Fact]
    public void Get_WithSessionId_ShouldReturnImageInfo()
    {
        var pptPath = CreateTestPresentation("test_session_get.pptx", addImages: true);
        var sessionId = OpenSession(pptPath);
        var result = _tool.Execute("get", sessionId: sessionId, slideIndex: 0);
        var data = GetResultData<GetImagesPptResult>(result);
        Assert.Equal(1, data.ImageCount);
    }

    [Fact]
    public void Add_WithSessionId_ShouldAddInMemory()
    {
        var pptPath = CreateTestPresentation("test_session_add.pptx");
        var imagePath = CreateTestImage("session_test_image.png");
        var sessionId = OpenSession(pptPath);
        var ppt = SessionManager.GetDocument<Presentation>(sessionId);
        var initialCount = ppt.Slides[0].Shapes.OfType<IPictureFrame>().Count();
        var result = _tool.Execute("add", sessionId: sessionId, slideIndex: 0, imagePath: imagePath, x: 100, y: 100,
            width: 200, height: 150);
        var data = GetResultData<SuccessResult>(result);
        Assert.StartsWith("Image added to slide", data.Message);
        Assert.True(ppt.Slides[0].Shapes.OfType<IPictureFrame>().Count() > initialCount);
    }

    [Fact]
    public void Edit_WithSessionId_ShouldEditInMemory()
    {
        var pptPath = CreateTestPresentation("test_session_edit.pptx", addImages: true);
        var sessionId = OpenSession(pptPath);
        var ppt = SessionManager.GetDocument<Presentation>(sessionId);
        var result = _tool.Execute("edit", sessionId: sessionId, slideIndex: 0, imageIndex: 0, width: 400, height: 300);
        var data = GetResultData<SuccessResult>(result);
        Assert.StartsWith("Image", data.Message);
        Assert.Contains("updated", data.Message);
        var image = ppt.Slides[0].Shapes.OfType<IPictureFrame>().First();
        Assert.Equal(400, image.Width);
    }

    [Fact]
    public void Delete_WithSessionId_ShouldDeleteInMemory()
    {
        var pptPath = CreateTestPresentation("test_session_delete.pptx", addImages: true);
        var sessionId = OpenSession(pptPath);
        var ppt = SessionManager.GetDocument<Presentation>(sessionId);
        var initialCount = ppt.Slides[0].Shapes.OfType<IPictureFrame>().Count();
        var result = _tool.Execute("delete", sessionId: sessionId, slideIndex: 0, imageIndex: 0);
        var data = GetResultData<SuccessResult>(result);
        Assert.StartsWith("Image", data.Message);
        Assert.Contains("deleted from slide", data.Message);
        Assert.True(ppt.Slides[0].Shapes.OfType<IPictureFrame>().Count() < initialCount);
    }

    [Fact]
    public void Execute_WithInvalidSessionId_ShouldThrowKeyNotFoundException()
    {
        Assert.Throws<KeyNotFoundException>(() => _tool.Execute("get", sessionId: "invalid_session", slideIndex: 0));
    }

    [Fact]
    public void Execute_WithBothPathAndSessionId_ShouldPreferSessionId()
    {
        var pptPath1 = CreateTestPresentation("test_path_image.pptx");
        var pptPath2 = CreateTestPresentation("test_session_image.pptx", addImages: true);
        var sessionId = OpenSession(pptPath2);
        var result = _tool.Execute("get", pptPath1, sessionId, slideIndex: 0);
        var data = GetResultData<GetImagesPptResult>(result);
        Assert.Equal(1, data.ImageCount);
    }

    #endregion
}
