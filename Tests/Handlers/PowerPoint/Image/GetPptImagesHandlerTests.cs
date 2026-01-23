using System.Drawing;
using System.Drawing.Imaging;
using Aspose.Slides;
using AsposeMcpServer.Handlers.PowerPoint.Image;
using AsposeMcpServer.Results.PowerPoint.Image;
using AsposeMcpServer.Tests.Infrastructure;

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

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetImagesPptResult>(res);

        Assert.True(result.ImageCount >= 0);
        Assert.NotNull(result.Images);
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

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetImagesPptResult>(res);

        Assert.Equal(0, result.ImageCount);
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

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetImagesPptResult>(res);

        Assert.True(result.Images.Count > 0);
        var firstImage = result.Images[0];
        Assert.True(firstImage.X >= 0);
        Assert.True(firstImage.Y >= 0);
        Assert.True(firstImage.Width >= 0);
        Assert.True(firstImage.Height >= 0);
    }

    #endregion
}
