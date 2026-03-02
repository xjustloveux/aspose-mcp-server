using System.Runtime.Versioning;
using Aspose.Slides;
using AsposeMcpServer.Handlers.PowerPoint.Image;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.PowerPoint.Image;

[SupportedOSPlatform("windows")]
public class AddPptImageHandlerTests : PptHandlerTestBase
{
    private readonly AddPptImageHandler _handler = new();

    #region Operation Property

    [SkippableFact]
    public void Operation_Returns_Add()
    {
        SkipIfNotWindows();
        Assert.Equal("add", _handler.Operation);
    }

    #endregion

    #region Helper Methods

    private static List<PictureFrame> GetPictureFrames(ISlide slide)
    {
        return slide.Shapes.OfType<PictureFrame>().ToList();
    }

    #endregion

    #region Basic Add Operations

    [SkippableFact]
    public void Execute_AddsImageToSlide()
    {
        SkipIfNotWindows();
        var tempImagePath = CreateTempImageFile();
        var pres = CreateEmptyPresentation();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 0 },
            { "imagePath", tempImagePath }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
        if (!IsEvaluationMode())
        {
            var pictureFrames = GetPictureFrames(pres.Slides[0]);
            Assert.True(pictureFrames.Count > 0);
        }

        AssertModified(context);
    }

    [SkippableFact]
    public void Execute_WithCustomPosition_AddsImageAtPosition()
    {
        SkipIfNotWindows();
        var tempImagePath = CreateTempImageFile();
        var pres = CreateEmptyPresentation();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 0 },
            { "imagePath", tempImagePath },
            { "x", 150f },
            { "y", 200f }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
        if (!IsEvaluationMode())
        {
            var pictureFrames = GetPictureFrames(pres.Slides[0]);
            Assert.True(pictureFrames.Count > 0);
            Assert.Equal(150f, pictureFrames[0].X);
            Assert.Equal(200f, pictureFrames[0].Y);
        }

        AssertModified(context);
    }

    [SkippableFact]
    public void Execute_WithCustomSize_AddsImageWithSize()
    {
        SkipIfNotWindows();
        var tempImagePath = CreateTempImageFile();
        var pres = CreateEmptyPresentation();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 0 },
            { "imagePath", tempImagePath },
            { "width", 300f },
            { "height", 200f }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
        if (!IsEvaluationMode())
        {
            var pictureFrames = GetPictureFrames(pres.Slides[0]);
            Assert.True(pictureFrames.Count > 0);
            Assert.Equal(300f, pictureFrames[0].Width);
            Assert.Equal(200f, pictureFrames[0].Height);
        }

        AssertModified(context);
    }

    #endregion

    #region Error Handling

    [SkippableFact]
    public void Execute_WithNonExistentImage_ThrowsFileNotFoundException()
    {
        SkipIfNotWindows();
        var pres = CreateEmptyPresentation();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 0 },
            { "imagePath", "non_existent_image.png" }
        });

        Assert.Throws<FileNotFoundException>(() => _handler.Execute(context, parameters));
    }

    [SkippableFact]
    public void Execute_WithInvalidSlideIndex_ThrowsException()
    {
        SkipIfNotWindows();
        var tempImagePath = CreateTempImageFile();
        var pres = CreateEmptyPresentation();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 99 },
            { "imagePath", tempImagePath }
        });

        Assert.ThrowsAny<Exception>(() => _handler.Execute(context, parameters));
    }

    #endregion
}
