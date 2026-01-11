using Aspose.Slides;
using AsposeMcpServer.Handlers.PowerPoint.Image;
using AsposeMcpServer.Tests.Helpers;

namespace AsposeMcpServer.Tests.Handlers.PowerPoint.Image;

public class AddPptImageHandlerTests : PptHandlerTestBase
{
    private readonly AddPptImageHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_Add()
    {
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

    [Fact]
    public void Execute_AddsImageToSlide()
    {
        var tempImagePath = CreateTempImageFile();
        var pres = CreateEmptyPresentation();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 0 },
            { "imagePath", tempImagePath }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("added", result.ToLower());
        var pictureFrames = GetPictureFrames(pres.Slides[0]);
        Assert.True(pictureFrames.Count > 0);
        AssertModified(context);
    }

    [Fact]
    public void Execute_WithCustomPosition_AddsImageAtPosition()
    {
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

        var result = _handler.Execute(context, parameters);

        Assert.Contains("added", result.ToLower());
        var pictureFrames = GetPictureFrames(pres.Slides[0]);
        Assert.True(pictureFrames.Count > 0);
        Assert.Equal(150f, pictureFrames[0].X);
        Assert.Equal(200f, pictureFrames[0].Y);
    }

    [Fact]
    public void Execute_WithCustomSize_AddsImageWithSize()
    {
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

        var result = _handler.Execute(context, parameters);

        Assert.Contains("added", result.ToLower());
        var pictureFrames = GetPictureFrames(pres.Slides[0]);
        Assert.True(pictureFrames.Count > 0);
        Assert.Equal(300f, pictureFrames[0].Width);
        Assert.Equal(200f, pictureFrames[0].Height);
    }

    #endregion

    #region Error Handling

    [Fact]
    public void Execute_WithNonExistentImage_ThrowsFileNotFoundException()
    {
        var pres = CreateEmptyPresentation();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 0 },
            { "imagePath", "non_existent_image.png" }
        });

        Assert.Throws<FileNotFoundException>(() => _handler.Execute(context, parameters));
    }

    [Fact]
    public void Execute_WithInvalidSlideIndex_ThrowsException()
    {
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
