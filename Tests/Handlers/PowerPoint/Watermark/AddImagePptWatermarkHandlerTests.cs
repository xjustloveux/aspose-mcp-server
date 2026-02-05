using AsposeMcpServer.Handlers.PowerPoint.Watermark;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.PowerPoint.Watermark;

public class AddImagePptWatermarkHandlerTests : PptHandlerTestBase
{
    private readonly AddImagePptWatermarkHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_AddImage()
    {
        Assert.Equal("add_image", _handler.Operation);
    }

    #endregion

    /// <summary>
    ///     Creates a minimal test image file (1x1 white pixel PNG) with a relative path
    ///     to comply with SecurityHelper.ValidateFilePath which rejects absolute paths.
    /// </summary>
    /// <returns>The relative file name of the created image file.</returns>
    private string CreateTestImage()
    {
        var fileName = $"test_wm_{Guid.NewGuid()}.png";
        var pngBytes = new byte[]
        {
            0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A, // PNG signature
            0x00, 0x00, 0x00, 0x0D, 0x49, 0x48, 0x44, 0x52, // IHDR chunk
            0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, // 1x1
            0x08, 0x02, 0x00, 0x00, 0x00, 0x90, 0x77, 0x53, 0xDE, // 8-bit RGB
            0x00, 0x00, 0x00, 0x0C, 0x49, 0x44, 0x41, 0x54, // IDAT chunk
            0x08, 0xD7, 0x63, 0xF8, 0xCF, 0xC0, 0x00, 0x00,
            0x00, 0x02, 0x00, 0x01, 0xE2, 0x21, 0xBC, 0x33,
            0x00, 0x00, 0x00, 0x00, 0x49, 0x45, 0x4E, 0x44, // IEND chunk
            0xAE, 0x42, 0x60, 0x82
        };
        File.WriteAllBytes(fileName, pngBytes);
        TestFiles.Add(Path.GetFullPath(fileName));
        return fileName;
    }

    #region Add Image Watermark

    [Fact]
    public void Execute_WithValidImage_ReturnsSuccessResult()
    {
        var pres = CreateEmptyPresentation();
        var context = CreateContext(pres);
        var imagePath = CreateTestImage();
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "imagePath", imagePath }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);
        Assert.Contains("1 slide(s)", result.Message);
    }

    [Fact]
    public void Execute_WithValidImage_MarksContextModified()
    {
        var pres = CreateEmptyPresentation();
        var context = CreateContext(pres);
        var imagePath = CreateTestImage();
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "imagePath", imagePath }
        });

        _handler.Execute(context, parameters);

        AssertModified(context);
    }

    [Fact]
    public void Execute_WithMultipleSlides_AddsToAllSlides()
    {
        var pres = CreatePresentationWithSlides(3);
        var context = CreateContext(pres);
        var imagePath = CreateTestImage();
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "imagePath", imagePath }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);
        Assert.Contains("3 slide(s)", result.Message);
    }

    [Fact]
    public void Execute_WithMissingImagePath_ThrowsArgumentException()
    {
        var pres = CreateEmptyPresentation();
        var context = CreateContext(pres);
        var parameters = CreateEmptyParameters();

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    [Fact]
    public void Execute_WithNonExistentFile_ThrowsArgumentException()
    {
        var pres = CreateEmptyPresentation();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "imagePath", "nonexistent_watermark_image.png" }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    [Fact]
    public void Execute_WatermarkShapeNameContainsImagePrefix()
    {
        var pres = CreateEmptyPresentation();
        var context = CreateContext(pres);
        var imagePath = CreateTestImage();
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "imagePath", imagePath }
        });

        _handler.Execute(context, parameters);

        var shapes = pres.Slides[0].Shapes;
        var hasImageWatermark = false;
        foreach (var shape in shapes)
            if (shape.Name != null && shape.Name.Contains("_IMAGE_"))
            {
                hasImageWatermark = true;
                break;
            }

        Assert.True(hasImageWatermark);
    }

    [Fact]
    public void Execute_WithCustomDimensions_ReturnsSuccessResult()
    {
        var pres = CreateEmptyPresentation();
        var context = CreateContext(pres);
        var imagePath = CreateTestImage();
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "imagePath", imagePath },
            { "width", 300f },
            { "height", 300f }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
    }

    #endregion
}
