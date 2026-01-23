using System.Drawing;
using System.Drawing.Imaging;
using Aspose.Slides;
using AsposeMcpServer.Handlers.PowerPoint.Image;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;

#pragma warning disable CA1416

namespace AsposeMcpServer.Tests.Handlers.PowerPoint.Image;

public class DeletePptImageHandlerTests : PptHandlerTestBase
{
    private readonly DeletePptImageHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_Delete()
    {
        Assert.Equal("delete", _handler.Operation);
    }

    #endregion

    #region Basic Delete Operations

    [Fact]
    public void Execute_DeletesImage()
    {
        var pres = CreatePresentationWithImage();
        var initialCount = GetPictureFrames(pres.Slides[0]).Count;
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 0 },
            { "imageIndex", 0 }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("deleted", result.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Equal(initialCount - 1, GetPictureFrames(pres.Slides[0]).Count);
        AssertModified(context);
    }

    #endregion

    #region Error Handling

    [Fact]
    public void Execute_WithInvalidImageIndex_ThrowsException()
    {
        var pres = CreatePresentationWithImage();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 0 },
            { "imageIndex", 99 }
        });

        Assert.ThrowsAny<Exception>(() => _handler.Execute(context, parameters));
    }

    [Fact]
    public void Execute_WithInvalidSlideIndex_ThrowsException()
    {
        var pres = CreatePresentationWithImage();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 99 },
            { "imageIndex", 0 }
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
        g.Clear(Color.Red);

        using var ms = new MemoryStream();
        bmp.Save(ms, ImageFormat.Bmp);
        ms.Position = 0;

        var image = pres.Images.AddImage(ms);
        slide.Shapes.AddPictureFrame(ShapeType.Rectangle, 50, 50, 100, 100, image);

        return pres;
    }

    private static List<PictureFrame> GetPictureFrames(ISlide slide)
    {
        return slide.Shapes.OfType<PictureFrame>().ToList();
    }

    #endregion
}
