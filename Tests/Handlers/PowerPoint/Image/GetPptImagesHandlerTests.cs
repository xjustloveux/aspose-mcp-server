using System.Drawing;
using System.Drawing.Imaging;
using System.Text.Json;
using Aspose.Slides;
using AsposeMcpServer.Handlers.PowerPoint.Image;
using AsposeMcpServer.Tests.Helpers;

#pragma warning disable CA1416

namespace AsposeMcpServer.Tests.Handlers.PowerPoint.Image;

public class GetPptImagesHandlerTests : PptHandlerTestBase
{
    private readonly GetPptImagesHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_Get()
    {
        Assert.Equal("get", _handler.Operation);
    }

    #endregion

    #region Read-Only Verification

    [Fact]
    public void Execute_DoesNotModifyPresentation()
    {
        var pres = CreatePresentationWithImage();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 0 }
        });

        _handler.Execute(context, parameters);

        AssertNotModified(context);
    }

    #endregion

    #region Error Handling

    [Fact]
    public void Execute_WithInvalidSlideIndex_ThrowsException()
    {
        var pres = CreateEmptyPresentation();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 99 }
        });

        Assert.ThrowsAny<Exception>(() => _handler.Execute(context, parameters));
    }

    #endregion

    #region Helper Methods

    private static Presentation CreatePresentationWithImage()
    {
        var pres = new Presentation();
        var slide = pres.Slides[0];

        using var bmp = new Bitmap(100, 100);
        using var g = Graphics.FromImage(bmp);
        g.Clear(Color.Green);

        using var ms = new MemoryStream();
        bmp.Save(ms, ImageFormat.Bmp);
        ms.Position = 0;

        var image = pres.Images.AddImage(ms);
        slide.Shapes.AddPictureFrame(ShapeType.Rectangle, 50, 50, 100, 100, image);

        return pres;
    }

    #endregion

    #region Basic Get Operations

    [Fact]
    public void Execute_ReturnsImageInfo()
    {
        var pres = CreatePresentationWithImage();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 0 }
        });

        var result = _handler.Execute(context, parameters);

        var json = JsonDocument.Parse(result);
        Assert.True(json.RootElement.TryGetProperty("imageCount", out _));
        Assert.True(json.RootElement.TryGetProperty("images", out _));
    }

    [Fact]
    public void Execute_WithEmptySlide_ReturnsZeroImages()
    {
        var pres = CreateEmptyPresentation();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 0 }
        });

        var result = _handler.Execute(context, parameters);

        var json = JsonDocument.Parse(result);
        Assert.Equal(0, json.RootElement.GetProperty("imageCount").GetInt32());
    }

    [Fact]
    public void Execute_ReturnsImagePosition()
    {
        var pres = CreatePresentationWithImage();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 0 }
        });

        var result = _handler.Execute(context, parameters);

        var json = JsonDocument.Parse(result);
        var images = json.RootElement.GetProperty("images");
        Assert.True(images.GetArrayLength() > 0);
        var firstImage = images[0];
        Assert.True(firstImage.TryGetProperty("x", out _));
        Assert.True(firstImage.TryGetProperty("y", out _));
        Assert.True(firstImage.TryGetProperty("width", out _));
        Assert.True(firstImage.TryGetProperty("height", out _));
    }

    #endregion
}
