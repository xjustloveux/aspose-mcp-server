using System.Drawing;
using System.Drawing.Imaging;
using Aspose.Slides;
using AsposeMcpServer.Handlers.PowerPoint.Image;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;

// CA1416 - System.Drawing.Common is Windows-only, cross-platform support not required
#pragma warning disable CA1416

namespace AsposeMcpServer.Tests.Handlers.PowerPoint.Image;

public class EditPptImageHandlerTests : PptHandlerTestBase
{
    private readonly EditPptImageHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_Edit()
    {
        Assert.Equal("edit", _handler.Operation);
    }

    #endregion

    #region Basic Edit Operations

    [Fact]
    public void Execute_EditsImagePosition()
    {
        var pres = CreatePresentationWithImage();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 0 },
            { "imageIndex", 0 },
            { "x", 200f },
            { "y", 300f }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
        if (!IsEvaluationMode())
        {
            var pictureFrame = GetPictureFrames(pres.Slides[0])[0];
            Assert.Equal(200f, pictureFrame.X);
            Assert.Equal(300f, pictureFrame.Y);
        }

        AssertModified(context);
    }

    [Fact]
    public void Execute_EditsImageSize()
    {
        var pres = CreatePresentationWithImage();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 0 },
            { "imageIndex", 0 },
            { "width", 250f },
            { "height", 150f }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
        if (!IsEvaluationMode())
        {
            var pictureFrame = GetPictureFrames(pres.Slides[0])[0];
            Assert.Equal(250f, pictureFrame.Width);
            Assert.Equal(150f, pictureFrame.Height);
        }

        AssertModified(context);
    }

    [Fact]
    public void Execute_WithNewImage_ReplacesImage()
    {
        var tempImagePath = CreateTempImageFile();
        var pres = CreatePresentationWithImage();
        var originalImage = GetPictureFrames(pres.Slides[0])[0].PictureFormat.Picture.Image;
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 0 },
            { "imageIndex", 0 },
            { "imagePath", tempImagePath }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
        if (!IsEvaluationMode())
        {
            var pictureFrame = GetPictureFrames(pres.Slides[0])[0];
            Assert.NotEqual(originalImage, pictureFrame.PictureFormat.Picture.Image);
        }

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
        g.Clear(Color.Yellow);

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
